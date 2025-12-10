# Planificateur dynamique des livraisons par plantation (Rendement annuel)

## Installation (Windows/Mac/Linux)

1. Installe Python 3.10+
2. Dans un terminal :
   ```bash
   pip install -r requirements.txt
   ```

## Lancer l'application
```bash
streamlit run app.py
```
Ensuite, ouvre le lien local affiché (http://localhost:8501).

## Utilisation
1. Charge ton Excel contenant les colonnes : `CODE_PLANTATION, NOM, RENDEMENT`.
2. Choisis l'année, le mois (ou toute l'année) et la fréquence (1 à 3).
3. (Optionnel) Active le quota et choisis la plantation + le kg à tracer.
4. Clique sur **Générer plan** : tu verras le plan et pourras **Télécharger** l'Excel.
