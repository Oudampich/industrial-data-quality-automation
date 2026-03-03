# -*- coding: utf-8 -*-
"""
Module: controle_data_saisie
Description: Analyse de fichiers Excel pour valider les dates, SIRET, montants financiers,
             détecter des caractères non valides dans le texte et identifier les doublons.
"""

import os
import re
import pandas as pd
from dateutil.parser import parse

# ============================================
# Fonctions de validation
# ============================================

def is_valid_date(value):
    """Vérifie si la valeur est une date valide ou vide."""
    try:
        if pd.isnull(value) or str(value).strip() == '':
            return True
        parse(str(value), dayfirst=True, fuzzy=False)
        return True
    except:
        return False

def is_valid_financial(value):
    """Vérifie si la valeur est un montant financier valide ou vide."""
    try:
        if pd.isnull(value) or str(value).strip() == '':
            return True
        float(str(value).replace("€", "").replace(",", ".").replace(" ", ""))
        return True
    except:
        return False

def is_valid_siret(value):
    """Vérifie si la valeur est un SIRET valide (14 chiffres) ou vide."""
    v = str(value).strip()
    if v == '' or pd.isnull(value):
        return True
    return v.isdigit() and len(v) == 14

# ============================================
# Analyse du fichier Excel
# ============================================

def analyze_excel(file_path):
    """Analyse les colonnes spécifiques et renvoie un dictionnaire des anomalies."""
    if not os.path.isfile(file_path):
        print(f"Fichier introuvable : {file_path}")
        return None

    try:
        df = pd.read_excel(file_path, dtype=str)
    except Exception as e:
        print(f"Erreur lecture Excel : {e}")
        return None

    # Colonnes spécifiques
    colonnes_date = [
        "Date de réception de la demande actuelle DC4",
        "Date envoi mail statut dossier non recevable\n(Respect process SGP)",
        "Date du traitement de la demande par le pôle administratif / Innoha",
        "Date identique pour demande de précisions au fournisseur et demande avis GFI",
        "Date limite initiale de réception de précisions ",
        "Date prévisionnelle recalée de remise des précisions du fournisseur",
        "Date de retour de l'avis du GFI",
        "Date envoi du rejet du DC4",
        "Date Pré-validation par mail/Validé par mail",
        "Date de validation",
        "Date limite avant la tacite notification (vision FRNS)",
        "Date de notification",
        "Date prévisionnelle d'achèvement des prestations sous-traitées"
    ]

    colonnes_siret = [
        "N° SIRET associé",
        "Numéro SIRET",
        "Numéro SIRET "
    ]

    colonnes_financieres = [
        "Montant HT DC4"
    ]

    colonnes_texte = [col for col in df.columns if col not in colonnes_date + colonnes_siret + colonnes_financieres]

    results = {
        'dates_invalid': 0,
        'dates_invalid_lines': set(),
        'financial_invalid': 0,
        'financial_invalid_lines': set(),
        'siret_invalid': 0,
        'siret_invalid_lines': set(),
        'text_possible_issues': 0,
        'text_possible_issues_lines': set()
    }

    # Vérification dates
    for col in colonnes_date:
        if col in df.columns:
            for idx, val in df[col].items():
                if not is_valid_date(val):
                    results['dates_invalid'] += 1
                    results['dates_invalid_lines'].add(idx)

    # Vérification financières
    for col in colonnes_financieres:
        if col in df.columns:
            for idx, val in df[col].items():
                if not is_valid_financial(val):
                    results['financial_invalid'] += 1
                    results['financial_invalid_lines'].add(idx)

    # Vérification SIRET
    for col in colonnes_siret:
        if col in df.columns:
            for idx, val in df[col].items():
                if not is_valid_siret(val):
                    results['siret_invalid'] += 1
                    results['siret_invalid_lines'].add(idx)

    # Vérification texte
    for col in colonnes_texte:
        for idx, val in df[col].items():
            if pd.notnull(val) and re.search(r'[^a-zA-Z0-9À-ÿ\s\-\.,]', str(val)):
                results['text_possible_issues'] += 1
                results['text_possible_issues_lines'].add(idx)

    # Convertir sets en listes triées
    for key in results:
        if isinstance(results[key], set):
            results[key] = sorted(results[key])

    return results, df

# ============================================
# Détection doublons
# ============================================

def detect_duplicates(df, col_index=0):
    """Détecte les doublons dans la première colonne (hors vides)."""
    if df is None or df.empty:
        return None
    col_a = df.columns[col_index]
    df_clean = df[df[col_a].notna() & (df[col_a] != '')]
    duplicates = df_clean[df_clean.duplicated(subset=[col_a], keep=False)]
    return col_a, duplicates

# ============================================
# Main
# ============================================

def main():
    fichier = input("Chemin du fichier Excel à analyser : ")
    anomalies, df = analyze_excel(fichier)

    if anomalies is None:
        return

    print("\n=== Résumé des anomalies détectées ===")
    for k, v in anomalies.items():
        if 'lines' in k:
            print(f"{k} (lignes) : {v}")
        else:
            print(f"{k} (nombre) : {v}")

    col_a, doublons = detect_duplicates(df)
    if doublons is not None:
        print(f"\nNombre de lignes en doublon dans la colonne '{col_a}' : {len(doublons)}")

if __name__ == "__main__":
    main()
