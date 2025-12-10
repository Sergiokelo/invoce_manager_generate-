#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
from pathlib import Path
import pandas as pd
import numpy as np

JOIN_COLS = ["NOM", "AXE", "VILLAGE", "SUPERFICIE", "HECTARE"]
LEFT_CODE_COL = "CODE"
RIGHT_CODE_COL = "CODE_PLANTATION"

def read_excel_interactive(label_file, label_sheet):
    while True:
        path = input(f"{label_file} (chemin .xlsx): ").strip().strip('"').strip("'")
        if not path:
            print("Chemin vide, réessayez.")
            continue
        if not os.path.exists(path):
            print("Fichier introuvable, réessayez.")
            continue
        try:
            xl = pd.ExcelFile(path)
            print("Onglets disponibles:", ", ".join(xl.sheet_names))
            sheet = input(f"{label_sheet} (laisser vide pour 1er onglet): ").strip()
            if not sheet:
                sheet = xl.sheet_names[0]
            if sheet not in xl.sheet_names:
                print("Onglet introuvable, réessayez.")
                continue
            df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
            return df, path, sheet
        except Exception as e:
            print("Erreur lecture:", e)

def ensure_cols(df, cols, where):
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(f"Colonnes manquantes dans {where}: {missing}")

def main():
    print("="*68)
    print(" Remplacement strict des CODE du 1er fichier à partir du 2e fichier ")
    print(" Correspondance sur: NOM | AXE | VILLAGE | SUPERFICIE | HECTARE ")
    print(" Sortie: mêmes colonnes/ordre que le 1er fichier, CODE remplacé ")
    print("="*68)

    # 1) Lire les deux fichiers
    left_df, left_path, _ = read_excel_interactive(
        "Chemin du 1er fichier (MCK_DOUBLONS)",
        "Nom d’onglet du 1er fichier"
    )
    right_df, right_path, _ = read_excel_interactive(
        "Chemin du 2e fichier (TABLE MAP UPLOAD MCK)",
        "Nom d’onglet du 2e fichier"
    )

    # 2) Vérifs colonnes
    ensure_cols(left_df, JOIN_COLS + [LEFT_CODE_COL], "1er fichier")
    ensure_cols(right_df, JOIN_COLS + [RIGHT_CODE_COL], "2e fichier")

    # 3) Mémoriser l’ordre exact des colonnes du 1er fichier
    left_cols_order = list(left_df.columns)

    # 4) Dédupliquer le 2e fichier sur la clé d’appariement (garder 1er)
    right_unique = right_df.drop_duplicates(subset=JOIN_COLS, keep="first")

    # 5) Fusion (LEFT JOIN) pour récupérer le CODE_PLANTATION en face
    merged = left_df.merge(
        right_unique[JOIN_COLS + [RIGHT_CODE_COL]].rename(
            columns={RIGHT_CODE_COL: "__CODE_FROM_MAP__"}
        ),
        on=JOIN_COLS,
        how="left",
        copy=False
    )

    # 6) Remplacer CODE uniquement quand on a une correspondance
    has_new = merged["__CODE_FROM_MAP__"].notna()
    updated_count = int(has_new.sum())
    merged[LEFT_CODE_COL] = np.where(
        has_new, merged["__CODE_FROM_MAP__"], merged[LEFT_CODE_COL]
    )

    # 7) Retirer la colonne temporaire et REVENIR EXACTEMENT à la structure du 1er fichier
    if "__CODE_FROM_MAP__" in merged.columns:
        merged.drop(columns="__CODE_FROM_MAP__", inplace=True)

    # Conserver strictement l’ordre/structure d’origine
    output_df = merged[left_cols_order].copy()

    # 8) Sauvegarde
    default_out = str(Path(left_path).with_name(Path(left_path).stem + "_codes_replaced.xlsx"))
    print(f"\nChemin de sortie par défaut: {default_out}")
    out_path = input("Entrer un chemin de sortie (laisser vide pour défaut): ").strip()
    if not out_path:
        out_path = default_out

    out_dir = os.path.dirname(out_path)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    try:
        with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
            # un seul onglet, même structure que le 1er fichier
            output_df.to_excel(xw, index=False, sheet_name="MCK_DOUBLONS")
    except Exception as e:
        print("Erreur écriture:", e)
        sys.exit(1)

    # 9) Statut
    print("\n--- RÉSULTAT ---")
    print(f"Lignes dans le 1er fichier : {len(output_df)}")
    print(f"Codes remplacés (match trouvé): {updated_count}")
    print(f"Fichier écrit: {out_path}")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrompu.")
        sys.exit(130)
