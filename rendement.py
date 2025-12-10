import pandas as pd

# Utilisation de la chaîne brute pour le chemin d'accès
file_path = r'D:\doc\traçabilité\COPROCAB\registre.xlsx'

# Chargement du fichier en spécifiant le nom de la feuille
df = pd.read_excel(file_path, sheet_name="Registre Producteurs")

# Afficher les noms des colonnes pour comprendre la structure
print("Colonnes du fichier :", df.columns)

# Vérifier si la colonne cible est bien présente
if "Superficie" in df.columns:
    # Calcul de la récolte totale
    df["Recolte Totale"] = df["Superficie"] * 800
else:
    print("La colonne 'Superficie Totale de L'Exploitation (HA) *' est introuvable. Vérifiez le nom exact dans le fichier.")

# Sauvegarder dans un nouveau fichier
output_path = r'D:\doc\traçabilité\COPROCAB\COOPROCABE_Registre_planteurs_rempliX.xlsx'
df.to_excel(output_path, index=False)
print("Fichier mis à jour et sauvegardé.")



import pandas as pd

# Charger le fichier Excel
file_path = r"D:\doc\traçabilité\COPROCAM\DATACOPROCAM\TABLE D'ATTRIBUT COPROCAM.xlsx"  # Chemin de votre fichier source
data = pd.read_excel(file_path)

# Compter le nombre d'occurrences pour chaque code dans la colonne "CODE"
data['Nombre de plantation'] = data.groupby('CODE')['CODE'].transform('count')

# Supprimer les doublons et conserver les colonnes "CODE", "NOM", et "Nombre de plantation"
resultat = data[['CODE', 'NOM', 'Nombre de plantation']].drop_duplicates().reset_index(drop=True)

# Afficher le résultat dans la console
print(resultat)

# Enregistrer le résultat dans un nouveau fichier Excel
result_file_path = r"D:\doc\traçabilité\COPROCAM\DATACOPROCAM\nombre_plantations_par_code.xlsx"  # Chemin du fichier de sortie
resultat.to_excel(result_file_path, index=False)
