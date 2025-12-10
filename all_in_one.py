import pandas as pd

# === 1. Nettoyage initial du fichier source ===
# Charger le fichier d'origine
input_file_path = r"D:\doc\traçabilité\SOPAM\data_C3\TABLE D'ATTRIBUT SOPAM.xlsx"
cleaned_file_path = r"D:\doc\traçabilité\SOPAM\data_C3\TABLE_C3_P.xlsx"

df = pd.read_excel(input_file_path, sheet_name='Feuil1')

# Supprimer les suffixes "-P1", "-P2", etc. dans la colonne CODE
df['CODE'] = df['CODE'].str.replace(r'-P\d+', '', regex=True)

# Enregistrer ce fichier nettoyé (étape intermédiaire)
df.to_excel(cleaned_file_path, index=False)
print("Étape 1 terminée : Nettoyage et sauvegarde de MHN_data_C3.xlsx")

# === 2. Ajout de CODE_PLANTATION, RENDEMENT, CERTIFICATION ===
df = pd.read_excel(cleaned_file_path)

# Générer CODE_PLANTATION
df['Occurrence'] = df.groupby('CODE').cumcount() + 1
df['CODE_PLANTATION'] = df['CODE'] + '-P' + df['Occurrence'].astype(str)

# Calculer RENDEMENT = HECTARE * 900
df['RENDEMENT'] = df['HECTARE'] * 900

# Ajouter la colonne CERTIFICATION avec "NONE"
df['CERTIFICATION'] = "NONE"

# Supprimer les colonnes inutiles
df = df.drop(columns=['NOM', 'Occurrence'])

# Réorganiser les colonnes pour placer CODE_PLANTATION juste après CODE
cols = list(df.columns)
cols.insert(cols.index('CODE') + 1, cols.pop(cols.index('CODE_PLANTATION')))
df = df[cols]

# Enregistrer le fichier pour l’upload
upload_file_path = r"D:\doc\traçabilité\SOPAM\data_C3\DATA_C3_UPLOAD.xlsx"
df.to_excel(upload_file_path, index=False)
print("Étape 2 terminée : Sauvegarde du fichier préparé pour upload")

# === 3. Nombre de plantations par planteur ===
# Recharger le fichier intermédiaire nettoyé
df = pd.read_excel(cleaned_file_path)

# Calculer le nombre de plantations par CODE
df['Nombre de plantation'] = df.groupby('CODE')['CODE'].transform('count')

# Supprimer les doublons par CODE
resultat = df.drop_duplicates(subset='CODE').reset_index(drop=True)

# Conserver les colonnes nécessaires
resultat = resultat[['CODE', 'NOM', 'Nombre de plantation']]

# Enregistrer le résultat
result_path = r"D:\doc\traçabilité\SOPAM\data_C3\planteur x plantation.xlsx"
resultat.to_excel(result_path, index=False)
print("Étape 3 terminée : Sauvegarde du fichier planteur x plantation")

print("\n✅ Tous les fichiers ont été générés avec succès.")
