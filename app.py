import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import io
import os
import re

from openpyxl import load_workbook
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="GROUPE-e - Simulateur autoconsommation solaire", layout="wide")

LOGO_PATH = "logo.png"
PV_PROFILE_PATH = "profil_pv.csv"
PV_YIELD_KWH_PER_KWP = 1020.0  # rendement PV constant (kWh/kWp/an)

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


def validate_and_fix_curve(arr: np.ndarray, n_steps: int, label: str, check_negative: bool = True):
    """Complète par des 0 si des valeurs manquent, tronque si il y en a trop,
    et signale une erreur bloquante si des valeurs négatives sont présentes."""
    arr = np.array(arr, dtype=float)
    n = len(arr)
    if n < n_steps:
        arr = np.concatenate([arr, np.zeros(n_steps - n)])
        st.warning(f"⚠️ {label} : {n_steps - n} valeur(s) manquante(s) — complétée(s) par 0.")
    elif n > n_steps:
        arr = arr[:n_steps]
        st.warning(f"⚠️ {label} : {n - n_steps} valeur(s) en trop — tronquée(s).")

    if check_negative:
        neg_count = int(np.sum(arr < 0))
        if neg_count > 0:
            st.error(
                f"❌ {label} : {neg_count} valeur(s) négative(s) détectée(s). "
                "Une courbe de charge ne peut pas être négative — merci de corriger les données avant de continuer."
            )
            return None
    return arr


def kw_to_kwh_step(kw_array: np.ndarray) -> np.ndarray:
    """Convertit une courbe de charge en kW (puissance au quart d'heure) en énergie kWh
    pour ce pas de 15 min : kWh = kW x 0,25 h (= kW / 4)."""
    return kw_array * 0.25


def build_time_index(year: int) -> pd.DatetimeIndex:
    start = pd.Timestamp(year=year, month=1, day=1)
    end = pd.Timestamp(year=year + 1, month=1, day=1)
    return pd.date_range(start=start, end=end, freq='15min', inclusive='left')


@st.cache_data
def load_pv_profile():
    """Charge le profil PV unitaire (profil_pv.csv, une valeur par ligne, non normalisé à 1)."""
    if not os.path.exists(PV_PROFILE_PATH):
        return None
    try:
        arr = pd.read_csv(PV_PROFILE_PATH, header=None)[0].to_numpy(dtype=float)
        return arr
    except Exception:
        return None


def clear_field(key: str):
    st.session_state[key] = ""


def run_battery_simulation(prod_kwh: np.ndarray, conso_kwh: np.ndarray, capacity_kwh: float):
    """Simulation quart d'heure par quart d'heure : autoconsommation avant/après batterie.
    Limite de puissance de charge/décharge = 0.5 x capacité (C-rate 0.5). SOC borné [0, capacité]."""
    n = len(prod_kwh)
    autoconso_avant = np.minimum(prod_kwh, conso_kwh)
    surplus = prod_kwh - autoconso_avant
    deficit = conso_kwh - autoconso_avant

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


# ==========================================================
# Génération du rapport Excel (données + graphiques natifs)
# ==========================================================

def build_excel_report(df: pd.DataFrame, capacity_kwh: float, sim_year: int) -> bytes:
    # 1) Feuille "Simulation" : toutes les données brutes, écrites rapidement via pandas
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Simulation")
    buffer.seek(0)
    wb = load_workbook(buffer)

    ws_sim = wb["Simulation"]
    ws_sim.column_dimensions['A'].width = 20
    for cell in list(ws_sim['A'])[1:]:
        cell.number_format = 'DD/MM/YYYY HH:MM'
    for col_letter in ['B', 'C', 'D', 'E', 'F', 'G']:
        ws_sim.column_dimensions[col_letter].width = 16

    # 2) Indicateurs clés
    total_prod = float(df["Production_solaire_kWh"].sum())
    total_conso = float(df["Consommation_brute_kWh"].sum())
    total_auto_avant = float(df["Autoconsommation_avant_batterie_kWh"].sum())
    total_auto_apres = float(df["Autoconsommation_apres_batterie_kWh"].sum())
    taux_avant = (total_auto_avant / total_conso * 100) if total_conso > 0 else 0.0
    taux_apres = (total_auto_apres / total_conso * 100) if total_conso > 0 else 0.0

    # 3) Feuille cachée avec les données sources des graphiques
    ws_data = wb.create_sheet("DonnéesGraphiques")
    ws_data.sheet_state = "hidden"

    # -- Bloc mensuel (lignes 1-13) --
    monthly = df.set_index("Horodatage").resample("MS").sum(numeric_only=True)
    ws_data["A1"] = "Mois"
    ws_data["B1"] = "Autoconso. avant batterie (kWh)"
    ws_data["C1"] = "Autoconso. après batterie (kWh)"
    for i, (ts, row) in enumerate(monthly.iterrows(), start=2):
        ws_data.cell(row=i, column=1, value=ts.strftime("%b %Y"))
        ws_data.cell(row=i, column=2, value=float(row["Autoconsommation_avant_batterie_kWh"]))
        ws_data.cell(row=i, column=3, value=float(row["Autoconsommation_apres_batterie_kWh"]))
    month_last_row = 1 + len(monthly)

    # -- Bloc camembert AVANT batterie (lignes 16-18) --
    ws_data["A16"] = "Répartition avant batterie"
    ws_data["B16"] = "kWh"
    ws_data["A17"] = "Autoconsommé"
    ws_data["B17"] = total_auto_avant
    ws_data["A18"] = "Importé du réseau"
    ws_data["B18"] = max(total_conso - total_auto_avant, 0)

    # -- Bloc camembert APRÈS batterie (lignes 21-23) --
    ws_data["A21"] = "Répartition après batterie"
    ws_data["B21"] = "kWh"
    ws_data["A22"] = "Autoconsommé"
    ws_data["B22"] = total_auto_apres
    ws_data["A23"] = "Importé du réseau"
    ws_data["B23"] = max(total_conso - total_auto_apres, 0)

    # -- Bloc journée type (21 juin, lignes 26+) --
    try:
        day_date = pd.Timestamp(year=sim_year, month=6, day=21).date()
        day_df = df[df["Horodatage"].dt.date == day_date]
    except Exception:
        day_df = df.iloc[0:0]
    ws_data["A26"] = "Heure"
    ws_data["B26"] = "Production solaire (kWh)"
    ws_data["C26"] = "Consommation brute (kWh)"
    ws_data["D26"] = "Autoconso. avant batterie (kWh)"
    ws_data["E26"] = "Autoconso. après batterie (kWh)"
    day_row = 27
    for _, r in day_df.iterrows():
        ws_data.cell(row=day_row, column=1, value=r["Horodatage"].strftime("%H:%M"))
        ws_data.cell(row=day_row, column=2, value=float(r["Production_solaire_kWh"]))
        ws_data.cell(row=day_row, column=3, value=float(r["Consommation_brute_kWh"]))
        ws_data.cell(row=day_row, column=4, value=float(r["Autoconsommation_avant_batterie_kWh"]))
        ws_data.cell(row=day_row, column=5, value=float(r["Autoconsommation_apres_batterie_kWh"]))
        day_row += 1
    day_last_row = day_row - 1

    # 4) Feuille "Résumé" (placée en première position)
    ws = wb.create_sheet("Résumé", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 3
    for col_letter in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
        ws.column_dimensions[col_letter].width = 14

    accent = "1F6F5C"  # vert-bleu GROUPE-e (à ajuster à la charte exacte si besoin)
    light_fill = PatternFill(start_color="EAF4F1", end_color="EAF4F1", fill_type="solid")

    if os.path.exists(LOGO_PATH):
        try:
            img = XLImage(LOGO_PATH)
            img.width = 140
            img.height = 60
            ws.add_image(img, "B2")
        except Exception:
            pass

    ws["E2"] = "Rapport de simulation — Autoconsommation solaire"
    ws["E2"].font = Font(size=16, bold=True, color=accent)
    ws["E3"] = f"GROUPE-e · Simulation au quart d'heure · Année {sim_year}"
    ws["E3"].font = Font(size=11, italic=True, color="666666")

    ws["B7"] = "Indicateurs clés"
    ws["B7"].font = Font(size=13, bold=True, color=accent)

    kpi_rows = [
        ("Production solaire annuelle", f"{total_prod:,.0f} kWh".replace(",", " "), None),
        ("Consommation annuelle", f"{total_conso:,.0f} kWh".replace(",", " "), None),
        ("Autoconsommation avant batterie", f"{total_auto_avant:,.0f} kWh".replace(",", " "), f"{taux_avant:.1f} %"),
        ("Autoconsommation après batterie", f"{total_auto_apres:,.0f} kWh".replace(",", " "), f"{taux_apres:.1f} %"),
        ("Gain apporté par la batterie", f"{(total_auto_apres - total_auto_avant):,.0f} kWh".replace(",", " "), f"+{(taux_apres - taux_avant):.1f} pts"),
        ("Capacité batterie simulée", f"{capacity_kwh:,.0f} kWh".replace(",", " "), None),
    ]
    row0 = 9
    for i, (label, value, extra) in enumerate(kpi_rows):
        r = row0 + i
        ws.cell(row=r, column=2, value=label).font = Font(bold=True, size=11)
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="left")
        for c in range(2, 6):
            ws.cell(row=r, column=c).fill = light_fill
        val_cell = ws.cell(row=r, column=5, value=value)
        val_cell.font = Font(bold=True, size=13, color=accent)
        val_cell.alignment = Alignment(horizontal="right")
        if extra:
            ws.cell(row=r, column=6, value=extra).font = Font(bold=True, size=11, color="888888")
        ws.row_dimensions[r].height = 20

    # -- Graphique barres : autoconsommation mensuelle avant / après --
    bar = BarChart()
    bar.type = "col"
    bar.grouping = "clustered"
    bar.title = "Autoconsommation mensuelle — avant / après batterie"
    bar.y_axis.title = "kWh"
    bar.x_axis.title = "Mois"
    data_ref = Reference(ws_data, min_col=2, max_col=3, min_row=1, max_row=month_last_row)
    cats_ref = Reference(ws_data, min_col=1, min_row=2, max_row=month_last_row)
    bar.add_data(data_ref, titles_from_data=True)
    bar.set_categories(cats_ref)
    bar.width = 24
    bar.height = 10
    ws.add_chart(bar, "B18")

    # -- Camemberts avant / après --
    pie_avant = PieChart()
    pie_avant.title = "Répartition avant batterie"
    data_ref = Reference(ws_data, min_col=2, max_col=2, min_row=16, max_row=18)
    cats_ref = Reference(ws_data, min_col=1, min_row=17, max_row=18)
    pie_avant.add_data(data_ref, titles_from_data=True)
    pie_avant.set_categories(cats_ref)
    pie_avant.dataLabels = DataLabelList()
    pie_avant.dataLabels.showPercent = True
    pie_avant.width = 12
    pie_avant.height = 10
    ws.add_chart(pie_avant, "B39")

    pie_apres = PieChart()
    pie_apres.title = "Répartition après batterie"
    data_ref = Reference(ws_data, min_col=2, max_col=2, min_row=21, max_row=23)
    cats_ref = Reference(ws_data, min_col=1, min_row=22, max_row=23)
    pie_apres.add_data(data_ref, titles_from_data=True)
    pie_apres.set_categories(cats_ref)
    pie_apres.dataLabels = DataLabelList()
    pie_apres.dataLabels.showPercent = True
    pie_apres.width = 12
    pie_apres.height = 10
    ws.add_chart(pie_apres, "G39")

    # -- Profil d'une journée type (21 juin) --
    if day_last_row >= 27:
        line = LineChart()
        line.title = "Profil d'une journée type — 21 juin"
        line.y_axis.title = "kWh / 15 min"
        line.x_axis.title = "Heure"
        data_ref = Reference(ws_data, min_col=2, max_col=5, min_row=26, max_row=day_last_row)
        cats_ref = Reference(ws_data, min_col=1, min_row=27, max_row=day_last_row)
        line.add_data(data_ref, titles_from_data=True)
        line.set_categories(cats_ref)
        line.width = 24
        line.height = 10
        ws.add_chart(line, "B60")
        ws["B80"] = "💡 Dans l'application web, ce graphique est interactif : un sélecteur permet d'afficher n'importe quelle journée de l'année."
        ws["B80"].font = Font(italic=True, size=9, color="888888")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ==========================================================
# Interface
# ==========================================================

col_logo, col_title = st.columns([1, 5])
with col_logo:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=140)
with col_title:
    st.title("☀️ GROUPE-e — Simulateur d'autoconsommation solaire & batterie")
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
    col_a, col_b = st.columns([12, 1])
    with col_a:
        txt_prod = st.text_area("Production solaire (kW, courbe de charge au 1/4h)", height=150, key="prod_paste")
    with col_b:
        st.write("")
        st.write("")
        st.button("🗑️", key="clear_prod_btn", on_click=clear_field, args=("prod_paste",), help="Vider ce champ")
    if txt_prod:
        arr_kw = parse_pasted_numbers(txt_prod)
        arr_kw = validate_and_fix_curve(arr_kw, n_steps, "Production solaire (collée)")
        if arr_kw is not None:
            prod_kwh = kw_to_kwh_step(arr_kw)
            st.success(f"{n_steps} valeurs de production chargées et converties en kWh (kW × 0,25 h) ✅")
else:
    installed_kwc = st.number_input("Puissance installée (kWc)", min_value=0.0, value=10.0, step=0.5)
    st.caption(f"Rendement PV utilisé : {PV_YIELD_KWH_PER_KWP:.0f} kWh/kWp/an (constante), appliqué au profil `{PV_PROFILE_PATH}`.")
    pv_profile = load_pv_profile()
    if pv_profile is None:
        st.error(f"Fichier `{PV_PROFILE_PATH}` introuvable. Placez-le à côté de `app.py` (dans le dépôt GitHub).")
    else:
        pv_profile_fixed = validate_and_fix_curve(pv_profile, n_steps, f"Profil PV ({PV_PROFILE_PATH})", check_negative=False)
        if pv_profile_fixed is not None:
            # profil_pv.csv est exprimé en kWh/kWc au pas de temps ; on divise par 4 pour
            # ramener sa somme annuelle à 1, afin que production = kWc x 1020 kWh/kWp/an
            prod_kwh = installed_kwc * PV_YIELD_KWH_PER_KWP * (pv_profile_fixed / 4.0)
            st.info(f"Production annuelle estimée : {prod_kwh.sum():,.0f} kWh".replace(",", " "))

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
    col_a, col_b = st.columns([12, 1])
    with col_a:
        txt_conso = st.text_area("Consommation brute (kW, courbe de charge au 1/4h)", height=150, key="conso_paste")
    with col_b:
        st.write("")
        st.write("")
        st.button("🗑️", key="clear_conso_btn", on_click=clear_field, args=("conso_paste",), help="Vider ce champ")
    if txt_conso:
        arr_kw = parse_pasted_numbers(txt_conso)
        arr_kw = validate_and_fix_curve(arr_kw, n_steps, "Consommation brute (collée)")
        if arr_kw is not None:
            conso_brute_kwh = kw_to_kwh_step(arr_kw)
            st.success(f"{n_steps} valeurs de consommation chargées et converties en kWh (kW × 0,25 h) ✅")
else:
    st.write(
        "Collez le soutirage réseau (consommation nette actuelle, avec autoconsommation existante) "
        "et l'injection solaire, au quart d'heure, **en kW**. La consommation brute sera reconstituée automatiquement."
    )
    col1, col2 = st.columns(2)
    with col1:
        col_a, col_b = st.columns([12, 1])
        with col_a:
            txt_net = st.text_area("Soutirage réseau / consommation nette (kW, courbe de charge)", height=150, key="net_paste")
        with col_b:
            st.write("")
            st.write("")
            st.button("🗑️", key="clear_net_btn", on_click=clear_field, args=("net_paste",), help="Vider ce champ")
    with col2:
        col_a, col_b = st.columns([12, 1])
        with col_a:
            txt_inj = st.text_area("Injection solaire (kW, courbe de charge)", height=150, key="inj_paste")
        with col_b:
            st.write("")
            st.write("")
            st.button("🗑️", key="clear_inj_btn", on_click=clear_field, args=("inj_paste",), help="Vider ce champ")
    if txt_net and txt_inj:
        arr_net_kw = parse_pasted_numbers(txt_net)
        arr_inj_kw = parse_pasted_numbers(txt_inj)
        arr_net_kw = validate_and_fix_curve(arr_net_kw, n_steps, "Soutirage réseau (collé)")
        arr_inj_kw = validate_and_fix_curve(arr_inj_kw, n_steps, "Injection solaire (collée)")
        if arr_net_kw is not None and arr_inj_kw is not None:
            if prod_kwh is None:
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
        st.session_state["sim_capacity"] = float(capacity_kwh)

if not ready:
    st.info("Complétez les 3 étapes ci-dessus (production, consommation, batterie) pour activer la génération.")

if "result_df" in st.session_state:
    df = st.session_state["result_df"]
    sim_year = st.session_state.get("sim_year", int(year))
    sim_capacity = st.session_state.get("sim_capacity", float(capacity_kwh))
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

    # ---------- Répartition avant / après (aperçu) ----------
    st.subheader("🥧 Répartition avant / après batterie")
    pc1, pc2 = st.columns(2)
    with pc1:
        fig_pie_avant = go.Figure(data=[go.Pie(
            labels=["Autoconsommé", "Importé du réseau"],
            values=[total_auto_avant, max(total_conso - total_auto_avant, 0)],
            hole=0.4,
        )])
        fig_pie_avant.update_layout(title="Avant batterie")
        st.plotly_chart(fig_pie_avant, use_container_width=True)
    with pc2:
        fig_pie_apres = go.Figure(data=[go.Pie(
            labels=["Autoconsommé", "Importé du réseau"],
            values=[total_auto_apres, max(total_conso - total_auto_apres, 0)],
            hole=0.4,
        )])
        fig_pie_apres.update_layout(title="Après batterie")
        st.plotly_chart(fig_pie_apres, use_container_width=True)

    # ---------- Export ----------
    st.subheader("⬇️ Export Excel")
    excel_bytes = build_excel_report(df, sim_capacity, sim_year)
    st.download_button(
        "Télécharger le fichier Excel (avec graphiques)",
        data=excel_bytes,
        file_name=f"groupe-e_simulation_{sim_year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
