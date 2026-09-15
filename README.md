# Sunae — Simulateur d'autoconsommation solaire & batterie

Application Streamlit qui génère, au quart d'heure sur une année complète (jusqu'à ~35 040 lignes),
un fichier Excel avec : horodatage, production solaire, consommation brute, autoconsommation
avant/après batterie et consommation nette — plus des graphiques de comparaison avant/après batterie.

## Déploiement (GitHub + Streamlit Community Cloud)

1. Créez un dépôt GitHub et poussez-y ces 3 fichiers : `app.py`, `requirements.txt`, `README.md`.
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

- **Production solaire** : soit une courbe réelle collée (une valeur par ligne, au quart d'heure),
  soit un profil synthétique généré à partir d'une puissance installée (kWc) et d'un rendement
  spécifique cible (kWh/kWc/an) — approximation pour une latitude suisse, pas une donnée météo réelle.
- **Consommation** : soit la consommation brute totale collée directement, soit reconstituée à
  partir du soutirage réseau (compteur) et de l'injection solaire déjà mesurés
  (`conso_brute = soutirage + (production − injection)`).
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
