import pandas as pd

# Charger le fichier Excel
file_path = r"D:\doc\traçabilité\MNH\data_C3\MHN_data_C3.xlsx"  # Chemin de votre fichier source
data = pd.read_excel(file_path)

# Calculer le nombre de plantations pour chaque "CODE"
data['Nombre de plantation'] = data.groupby('CODE')['CODE'].transform('count')

# Supprimer les doublons en gardant une seule ligne par "CODE" (peu importe la valeur dans "NOM")
resultat = data.drop_duplicates(subset='CODE').reset_index(drop=True)

# Conserver uniquement les colonnes "CODE", "NOM", et "Nombre de plantation"
resultat = resultat[['CODE', 'NOM', 'Nombre de plantation']]

# Enregistrer le résultat dans un nouveau fichier Excel
result_file_path = r"D:\doc\traçabilité\MNH\data_C3\MHN_DATAC3_PlANTEUR X PLANTATION.xlsx"  # Chemin du fichier de sortie
resultat.to_excel(result_file_path, index=False)

# Afficher le résultat dans la console
print(resultat)
