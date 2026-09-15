import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import io
import re

st.set_page_config(page_title="Sunae - Simulateur autoconsommation solaire", layout="wide")

# ==========================================================
# Fonctions utilitaires
# ==========================================================

def parse_pasted_numbers(text: str) -> np.ndarray:
    """Transforme un texte collé depuis Excel (une valeur par ligne) en tableau de floats.
    Gère les décimales avec virgule (format européen) et les séparateurs de milliers."""
    if not text or not text.strip():
        return np.array([])
    tokens = re.split(r'[\n\r\t;]+', text.strip())
    values = []
    for tok in tokens:
        tok = tok.strip().replace('\xa0', '').replace(' ', '')
        if tok == '':
            continue
        if ',' in tok and '.' not in tok:
            tok = tok.replace(',', '.')
        else:
            tok = tok.replace(',', '')
        try:
            values.append(float(tok))
        except ValueError:
            continue
    return np.array(values, dtype=float)


def kw_to_kwh_step(kw_array: np.ndarray) -> np.ndarray:
    """Convertit une courbe de charge en kW (puissance au quart d'heure) en énergie kWh
    pour ce pas de 15 min : kWh = kW x 0,25 h (= kW / 4)."""
    return kw_array * 0.25


def build_time_index(year: int) -> pd.DatetimeIndex:
    """Index au quart d'heure, du 1er janvier 00:00 au 31 décembre 23:45 de l'année donnée."""
    start = pd.Timestamp(year=year, month=1, day=1)
    end = pd.Timestamp(year=year + 1, month=1, day=1)
    return pd.date_range(start=start, end=end, freq='15min', inclusive='left')


def generate_synthetic_solar(n_steps: int, index: pd.DatetimeIndex, installed_kwc: float,
                              specific_yield: float = 1000.0, seed: int = 42) -> np.ndarray:
    """Profil solaire synthétique au quart d'heure (approximation pour une latitude ~ Suisse).
    Ce n'est pas une donnée météo réelle : c'est une estimation raisonnable calibrée sur un
    rendement spécifique annuel cible (kWh/kWc/an)."""
    rng = np.random.default_rng(seed)
    day_of_year = index.dayofyear.values
    hour_decimal = index.hour.values + index.minute.values / 60.0

    # Durée du jour approximative (heures), latitude ~46.5°N
    daylength = 12 + 4.3 * np.sin(2 * np.pi * (day_of_year - 81) / 365)
    sunrise = 12 - daylength / 2
    sunset = 12 + daylength / 2

    # Facteur saisonnier d'irradiance (max ~solstice d'été, min ~solstice d'hiver)
    seasonal = 0.35 + 0.65 * (0.5 + 0.5 * np.cos(2 * np.pi * (day_of_year - 172) / 365))

    prod = np.zeros(n_steps)
    is_day = (hour_decimal > sunrise) & (hour_decimal < sunset)
    frac = np.clip((hour_decimal - sunrise) / np.maximum(sunset - sunrise, 1e-6), 0, 1)
    shape = np.sin(np.pi * frac) ** 1.3
    cloud_noise = rng.uniform(0.75, 1.0, size=n_steps)

    prod[is_day] = installed_kwc * shape[is_day] * seasonal[is_day] * cloud_noise[is_day]
    prod = np.clip(prod, 0, None)

    # Recalibrage pour atteindre le rendement spécifique visé (kWh/kWc/an)
    # prod représente ici directement l'énergie (kWh) produite sur chaque pas de 15 min
    target_total = specific_yield * installed_kwc
    current_total = prod.sum()
    if current_total > 0:
        prod = prod * (target_total / current_total)
    return prod  # kWh produits sur chaque pas de 15 min


def run_battery_simulation(prod_kwh: np.ndarray, conso_kwh: np.ndarray, capacity_kwh: float):
    """Simulation quart d'heure par quart d'heure : autoconsommation avant/après batterie.
    Limite de puissance de charge/décharge = 0.5 x capacité (C-rate 0.5). SOC borné [0, capacité]."""
    n = len(prod_kwh)
    autoconso_avant = np.minimum(prod_kwh, conso_kwh)
    surplus = prod_kwh - autoconso_avant   # solaire non consommé directement
    deficit = conso_kwh - autoconso_avant  # besoin non couvert par le solaire direct

    charge = np.zeros(n)
    discharge = np.zeros(n)
    soc = np.zeros(n)

    power_limit_kw = 0.5 * capacity_kwh
    max_step_energy = power_limit_kw * 0.25  # kWh par pas de 15 min

    soc_kwh = 0.0
    for i in range(n):
        if capacity_kwh > 0 and surplus[i] > 0 and soc_kwh < capacity_kwh:
            e = min(surplus[i], max_step_energy, capacity_kwh - soc_kwh)
            soc_kwh += e
            charge[i] = e
        elif capacity_kwh > 0 and deficit[i] > 0 and soc_kwh > 0:
            e = min(deficit[i], max_step_energy, soc_kwh)
            soc_kwh -= e
            discharge[i] = e
        soc[i] = soc_kwh

    autoconso_apres = autoconso_avant + discharge
    conso_nette = conso_kwh - autoconso_apres
    soc_pct = (soc / capacity_kwh * 100) if capacity_kwh > 0 else np.zeros(n)
    return autoconso_avant, autoconso_apres, conso_nette, soc, soc_pct


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Simulation')
    return buffer.getvalue()


# ==========================================================
# Interface
# ==========================================================

st.title("☀️ Sunae — Simulateur d'autoconsommation solaire & batterie")
st.caption(
    "Génère un fichier Excel au quart d'heure sur une année complète : production solaire, "
    "consommation, autoconsommation avant/après batterie et consommation nette."
)

year = st.number_input("Année de simulation (pour l'horodatage)", min_value=2000, max_value=2100, value=2026, step=1)
index = build_time_index(int(year))
n_steps = len(index)
st.caption(f"Nombre de pas de temps attendu pour {year} : **{n_steps:,}**".replace(",", " "))

st.divider()

# ---------- Étape 1 : Production solaire ----------
st.header("1️⃣ Production solaire")
mode_prod = st.selectbox(
    "Source des données de production solaire",
    ["Courbe de charge réelle (copier-coller)", "Données virtuelles (estimation depuis puissance installée)"],
)

prod_kwh = None
if mode_prod.startswith("Courbe"):
    st.write(
        "Collez ici la colonne de la courbe de charge de production solaire au quart d'heure, "
        "**en kW** (une valeur par ligne, copiée depuis Excel)."
    )
    txt_prod = st.text_area("Production solaire (kW, courbe de charge au 1/4h)", height=150, key="prod_paste")
    if txt_prod:
        arr_kw = parse_pasted_numbers(txt_prod)
        if len(arr_kw) != n_steps:
            st.warning(f"{len(arr_kw)} valeurs détectées, {n_steps} attendues pour {year}. Vérifiez le collage.")
        else:
            prod_kwh = kw_to_kwh_step(arr_kw)
            st.success(f"{len(arr_kw)} valeurs de production chargées et converties en kWh (kW × 0,25 h) ✅")
else:
    installed_kwc = st.number_input("Puissance installée (kWc)", min_value=0.0, value=10.0, step=0.5)
    specific_yield = st.number_input("Rendement spécifique visé (kWh/kWc/an)", min_value=500, max_value=1500, value=1000, step=50)
    prod_kwh = generate_synthetic_solar(n_steps, index, installed_kwc, specific_yield)
    st.info(f"Profil synthétique généré — production annuelle estimée : {prod_kwh.sum():,.0f} kWh".replace(",", " "))

st.divider()

# ---------- Étape 2 : Consommation ----------
st.header("2️⃣ Consommation du client")
mode_conso = st.selectbox(
    "Type de données de consommation disponibles",
    ["Consommation brute totale (sans autoconsommation)", "Consommation nette (soutirage réseau) + injection solaire"],
)

conso_brute_kwh = None
if mode_conso.startswith("Consommation brute"):
    st.write("Collez ici la courbe de charge de consommation brute totale du client au quart d'heure, **en kW**.")
    txt_conso = st.text_area("Consommation brute (kW, courbe de charge au 1/4h)", height=150, key="conso_paste")
    if txt_conso:
        arr_kw = parse_pasted_numbers(txt_conso)
        if len(arr_kw) != n_steps:
            st.warning(f"{len(arr_kw)} valeurs détectées, {n_steps} attendues pour {year}.")
        else:
            conso_brute_kwh = kw_to_kwh_step(arr_kw)
            st.success(f"{len(arr_kw)} valeurs de consommation chargées et converties en kWh (kW × 0,25 h) ✅")
else:
    st.write(
        "Collez le soutirage réseau (consommation nette actuelle, avec autoconsommation existante) "
        "et l'injection solaire, au quart d'heure, **en kW**. La consommation brute sera reconstituée automatiquement."
    )
    col1, col2 = st.columns(2)
    with col1:
        txt_net = st.text_area("Soutirage réseau / consommation nette (kW, courbe de charge)", height=150, key="net_paste")
    with col2:
        txt_inj = st.text_area("Injection solaire (kW, courbe de charge)", height=150, key="inj_paste")
    if txt_net and txt_inj:
        arr_net_kw = parse_pasted_numbers(txt_net)
        arr_inj_kw = parse_pasted_numbers(txt_inj)
        if len(arr_net_kw) != n_steps or len(arr_inj_kw) != n_steps:
            st.warning(f"Soutirage : {len(arr_net_kw)} valeurs, Injection : {len(arr_inj_kw)} valeurs, {n_steps} attendues.")
        elif prod_kwh is None:
            st.warning("Renseignez d'abord la production solaire (étape 1).")
        else:
            net_kwh = kw_to_kwh_step(arr_net_kw)
            inj_kwh = kw_to_kwh_step(arr_inj_kw)
            # conso_brute = soutirage_reseau + autoconso_existante, avec autoconso_existante = production - injection
            conso_brute_kwh = net_kwh + (prod_kwh - inj_kwh)
            conso_brute_kwh = np.clip(conso_brute_kwh, 0, None)
            st.success("Consommation brute reconstituée (courbes converties kW → kWh) ✅")

st.divider()

# ---------- Étape 3 : Batterie ----------
st.header("3️⃣ Batterie à simuler")
capacity_kwh = st.slider("Capacité de la batterie (kWh)", min_value=0, max_value=300, value=10, step=1)
st.caption(f"Puissance de charge/décharge max simulée : {0.5 * capacity_kwh:.1f} kW (0,5 × capacité)")

st.divider()

# ---------- Génération ----------
ready = (
    prod_kwh is not None and conso_brute_kwh is not None
    and len(prod_kwh) == n_steps and len(conso_brute_kwh) == n_steps
)

if st.button("🚀 Générer la simulation", disabled=not ready, type="primary"):
    with st.spinner("Simulation quart d'heure par quart d'heure en cours..."):
        autoconso_avant, autoconso_apres, conso_nette, soc, soc_pct = run_battery_simulation(
            prod_kwh, conso_brute_kwh, float(capacity_kwh)
        )
        df = pd.DataFrame({
            "Horodatage": index,
            "Production_solaire_kWh": prod_kwh,
            "Consommation_brute_kWh": conso_brute_kwh,
            "Autoconsommation_avant_batterie_kWh": autoconso_avant,
            "Autoconsommation_apres_batterie_kWh": autoconso_apres,
            "Consommation_nette_kWh": conso_nette,
            "SOC_batterie_pct": soc_pct,
        })
        st.session_state["result_df"] = df
        st.session_state["sim_year"] = int(year)

if not ready:
    st.info("Complétez les 3 étapes ci-dessus (production, consommation, batterie) pour activer la génération.")

if "result_df" in st.session_state:
    df = st.session_state["result_df"]
    sim_year = st.session_state.get("sim_year", int(year))
    st.success("Simulation terminée ✅")

    # ---------- Indicateurs ----------
    total_prod = df["Production_solaire_kWh"].sum()
    total_conso = df["Consommation_brute_kWh"].sum()
    total_auto_avant = df["Autoconsommation_avant_batterie_kWh"].sum()
    total_auto_apres = df["Autoconsommation_apres_batterie_kWh"].sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Production annuelle", f"{total_prod:,.0f} kWh".replace(",", " "))
    c2.metric("Consommation annuelle", f"{total_conso:,.0f} kWh".replace(",", " "))
    if total_conso > 0:
        c3.metric("Taux d'autoconso. avant batterie", f"{total_auto_avant / total_conso * 100:.1f} %")
        c4.metric(
            "Taux d'autoconso. après batterie",
            f"{total_auto_apres / total_conso * 100:.1f} %",
            delta=f"+{(total_auto_apres - total_auto_avant) / total_conso * 100:.1f} pts",
        )

    # ---------- Graphique mensuel ----------
    st.subheader("📊 Autoconsommation mensuelle — avant / après batterie")
    monthly = df.set_index("Horodatage").resample("MS").sum(numeric_only=True)
    monthly.index = monthly.index.strftime("%b")
    fig_month = go.Figure()
    fig_month.add_bar(name="Avant batterie", x=monthly.index, y=monthly["Autoconsommation_avant_batterie_kWh"])
    fig_month.add_bar(name="Après batterie", x=monthly.index, y=monthly["Autoconsommation_apres_batterie_kWh"])
    fig_month.update_layout(barmode="group", yaxis_title="kWh")
    st.plotly_chart(fig_month, use_container_width=True)

    # ---------- Profil journalier type ----------
    st.subheader("📈 Profil d'une journée")
    default_day = pd.Timestamp(year=sim_year, month=6, day=21).date()
    chosen_day = st.date_input(
        "Choisir une journée à afficher",
        value=default_day,
        min_value=index[0].date(),
        max_value=index[-1].date(),
    )
    day_df = df[df["Horodatage"].dt.date == chosen_day]
    if not day_df.empty:
        fig_day = go.Figure()
        fig_day.add_scatter(x=day_df["Horodatage"], y=day_df["Production_solaire_kWh"], name="Production solaire", fill='tozeroy')
        fig_day.add_scatter(x=day_df["Horodatage"], y=day_df["Consommation_brute_kWh"], name="Consommation brute")
        fig_day.add_scatter(x=day_df["Horodatage"], y=day_df["Autoconsommation_avant_batterie_kWh"], name="Autoconso avant batterie", line=dict(dash="dot"))
        fig_day.add_scatter(x=day_df["Horodatage"], y=day_df["Autoconsommation_apres_batterie_kWh"], name="Autoconso après batterie", line=dict(dash="dash"))
        fig_day.update_layout(yaxis_title="kWh / 15 min")
        st.plotly_chart(fig_day, use_container_width=True)

        fig_soc = go.Figure()
        fig_soc.add_scatter(x=day_df["Horodatage"], y=day_df["SOC_batterie_pct"], name="État de charge batterie (%)")
        fig_soc.update_layout(yaxis_title="SOC (%)", yaxis_range=[0, 100])
        st.plotly_chart(fig_soc, use_container_width=True)

    # ---------- Export ----------
    st.subheader("⬇️ Export Excel")
    excel_bytes = to_excel_bytes(df)
    st.download_button(
        "Télécharger le fichier Excel",
        data=excel_bytes,
        file_name=f"sunae_simulation_{sim_year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
