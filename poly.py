import pandas as pd  # Assurez-vous que cette ligne est présente

# Charger le fichier
file_path = r"D:\doc\traçabilité\MNH\data_C3\MHN_data_C3.xlsx"
df = pd.read_excel(file_path)

# Ajouter la colonne "CODE_PLANTATION" en fonction de l'itération de "CODE_PLANTEUR"
df['Occurrence'] = df.groupby('CODE_PLANTEUR').cumcount() + 1
df['CODE_PLANTATION'] = df['CODE_PLANTEUR'] + '-P' + df['Occurrence'].astype(str)

# Calculer la colonne "RENDEMENT" en multipliant les valeurs de la colonne "HECTARE" par 800
df['RENDEMENT'] = df['HECTARE'] * 900

# Ajouter la colonne "CERTIFICATION" avec des valeurs "NONE"
df['CERTIFICATION'] = "NONE"

# Supprimer les colonnes "NOM" et la colonne temporaire "Occurrence"
df = df.drop(columns=['NOM', 'Occurrence'])

# Réorganiser les colonnes pour placer "CODE_PLANTATION" juste après "CODE_PLANTEUR"
cols = list(df.columns)
cols.insert(cols.index('CODE_PLANTEUR') + 1, cols.pop(cols.index('CODE_PLANTATION')))
df = df[cols]

# Enregistrer le fichier modifié
output_path = r"D:\doc\traçabilité\MNH\data_C3\MHN_Data_C3_UPLOAD.xlsx"
df.to_excel(output_path, index=False)

print("Fichier enregistré à :", output_path)
