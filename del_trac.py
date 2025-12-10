import pandas as pd

# Lis ton fichier
df = pd.read_excel("doc\traçabilité\COCOA SOURCE\MCK_DOUBLONS_codes_replaced.xlsx")  # adapte le nom
# Garde seulement les codes type MCK-0825-P1 (exactement MCK-4 chiffres-Pn)
mask = df["CODE"].str.fullmatch(r"MCK-\d{4}-P\d+", na=False)
df_keep = df[mask].copy()

# (optionnel) si tu veux aussi retirer explicitement un motif comme MAO-CEL
# df_keep = df[~df["CODE"].str.contains(r"-MAO-CEL-", na=False)]

# Sauvegarde
df_keep.to_excel("codes_filtrés.xlsx", index=False)
print(f"Lignes gardées: {len(df_keep)}")
