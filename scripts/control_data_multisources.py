"""
Module: control_data_multisources
Description: Compare deux fichiers Excel ligne par ligne et colonne par colonne,
             et exporte les différences dans un fichier Excel.
"""

import os
import pandas as pd

# Si vous utilisez Colab, décommentez cette ligne :
# from google.colab import files

def charger_excel(file_path):
    """Charge un fichier Excel et renvoie son contenu sous forme de DataFrame."""
    if not os.path.isfile(file_path):
        print(f"Fichier introuvable : {file_path}")
        return None
    try:
        return pd.read_excel(file_path)
    except Exception as e:
        print(f"Erreur lecture Excel pour {file_path} : {e}")
        return None

def comparer_excel(fichier1, fichier2, fichier_export):
    """
    Compare deux fichiers Excel cellule par cellule.
    Export les différences dans un nouveau fichier Excel.
    """
    df1 = charger_excel(fichier1)
    df2 = charger_excel(fichier2)

    if df1 is None or df2 is None:
        return None

    # Vérifier la compatibilité des dimensions
    if df1.shape != df2.shape:
        print("Les fichiers n'ont pas la même forme (différence de lignes ou de colonnes).")
        return None

    differences = []

    for i in range(df1.shape[0]):
        for j in range(df1.shape[1]):
            if df1.iloc[i, j] != df2.iloc[i, j]:
                differences.append({
                    'Ligne': i + 1,
                    'Colonne': j + 1,
                    'Valeur Fichier 1': df1.iloc[i, j],
                    'Valeur Fichier 2': df2.iloc[i, j]
                })

    if differences:
        df_diff = pd.DataFrame(differences)
        try:
            df_diff.to_excel(fichier_export, index=False)
            print(f"Les différences ont été exportées vers : {fichier_export}")
            return fichier_export
        except Exception as e:
            print(f"Erreur lors de l'exportation : {e}")
            return None
    else:
        print("Les fichiers sont identiques, aucune différence à exporter.")
        return None

def download_file_colab(file_path):
    """Télécharge le fichier via Colab (si Colab utilisé)."""
    # if os.path.exists(file_path):
    #     files.download(file_path)
    # else:
    #     print("Fichier introuvable pour téléchargement :", file_path)
    pass  # Pour exécution locale

def main():
    """Exemple d'utilisation"""
    fichier1 = input("Chemin du premier fichier Excel : ")
    fichier2 = input("Chemin du second fichier Excel : ")
    fichier_export = "differences.xlsx"

    export_file = comparer_excel(fichier1, fichier2, fichier_export)

    if export_file:
        print("Vérification de l'existence du fichier exporté...")
        if os.path.exists(export_file):
            print(f"Le fichier {export_file} est prêt.")
            # download_file_colab(export_file)  # Décommentez pour Colab
        else:
            print("Le fichier exporté n'a pas été trouvé.")

if __name__ == "__main__":
    main()
