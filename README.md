# GROUPE-e — Simulateur d'autoconsommation solaire & batterie

Application Streamlit qui génère, au quart d'heure sur une année complète (jusqu'à ~35 040 lignes),
un fichier Excel avec : horodatage, production solaire, consommation brute, autoconsommation
avant/après batterie et consommation nette — plus des graphiques natifs (barres, camemberts,
profil journalier) et les indicateurs clés directement dans l'onglet "Résumé" du fichier exporté.

## Fichiers à mettre dans le dépôt GitHub

- `app.py` — l'application
- `requirements.txt` — dépendances
- `profil_pv.csv` — profil PV unitaire (une valeur par ligne, 35 040 lignes) utilisé pour le mode
  "Données virtuelles" (production = puissance installée × 1020 kWh/kWp/an × profil)
- `logo.png` — logo affiché en haut de l'application et dans l'export Excel

## Déploiement (GitHub + Streamlit Community Cloud)

1. Créez un dépôt GitHub et poussez-y ces fichiers : `app.py`, `requirements.txt`, `README.md`,
   `profil_pv.csv`, `logo.png`.
   ```bash
   git init
   git add app.py requirements.txt README.md
   git commit -m "Sunae - simulateur autoconsommation"
   git branch -M main
   git remote add origin https://github.com/<votre-compte>/<votre-repo>.git
   git push -u origin main
   ```
2. Allez sur [streamlit.io/cloud](https://streamlit.io/cloud), connectez votre compte GitHub.
3. Cliquez sur **New app**, choisissez votre dépôt, la branche `main` et le fichier `app.py`.
4. Cliquez sur **Deploy**. L'app sera accessible sur une URL du type
   `https://<nom>.streamlit.app`.

## Utilisation en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Logique de simulation

- **Production solaire** : soit une courbe réelle collée en **kW** (convertie en kWh via × 0,25 h),
  soit un mode virtuel : `production (kWh) = puissance installée (kWp) × 1020 (kWh/kWp/an) × profil_pv.csv`.
- **Consommation** : soit la consommation brute totale collée directement (**kW**), soit reconstituée à
  partir du soutirage réseau (compteur) et de l'injection solaire déjà mesurés, tous deux en **kW**
  (`conso_brute = soutirage + (production − injection)`).
- **Données collées** : les courbes trop courtes sont complétées par des 0, les trop longues sont
  tronquées, et toute valeur négative bloque la génération avec un message d'erreur explicite.
- **Batterie** : capacité réglable de 0 à 300 kWh. Puissance de charge/décharge max = 0,5 × capacité
  (C-rate 0,5). SOC borné entre 0 % et 100 %. À chaque pas de 15 min : si la production dépasse la
  consommation, le surplus charge la batterie (dans la limite de la puissance et de la capacité
  restante) ; sinon, le déficit est comblé par décharge de la batterie (dans la limite de la
  puissance et de l'énergie disponible).

## Notes

- Le nombre de valeurs collées doit correspondre exactement au nombre de pas de 15 min de l'année
  choisie (35 040 pour une année normale, 35 136 pour une année bissextile) — l'app l'indique en
  haut de page.
- Le profil solaire virtuel est une estimation simplifiée à but d'illustration ; pour un
  dimensionnement précis, privilégiez une courbe réelle ou des données PVGIS.
