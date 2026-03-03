"""
Module: automation_mail
Description: Upload d'un fichier Excel, détection des valeurs manquantes (NA),
             création d'un rapport et envoi par email.
"""

import os
import io
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# Si vous utilisez Colab, décommentez les lignes suivantes :
# from google.colab import files

def upload_excel_file_colab():
    """Permet à l'utilisateur d'uploader un fichier Excel dans Colab."""
    # uploaded = files.upload()  # Décommentez pour Colab
    # if not uploaded:
    #     print("Aucun fichier uploadé !")
    #     return None
    # file_name = next(iter(uploaded))
    # return io.BytesIO(uploaded[file_name])

    # Pour test local, remplacer par le chemin de fichier
    file_path = input("Chemin du fichier Excel : ")
    if not os.path.isfile(file_path):
        print("Fichier introuvable !")
        return None
    return file_path

def detect_na_in_excel(file_path):
    """Lit le fichier Excel et retourne les lignes contenant des valeurs manquantes."""
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Erreur lecture Excel : {e}")
        return None

    na_rows = df[df.isna().any(axis=1)]
    if not na_rows.empty:
        print(f"{len(na_rows)} ligne(s) avec NA détectées")
        return na_rows
    else:
        print("Aucun NA trouvé")
        return None

def create_na_report(na_data, output_filename="NA_Report.xlsx"):
    """Crée un fichier Excel avec les lignes contenant des NA."""
    if na_data is not None and not na_data.empty:
        na_data.to_excel(output_filename, index=False)
        print(f"Fichier {output_filename} créé")
        return output_filename
    return None

def send_na_report(email_recipient, attachment_path, email_sender, email_password):
    """Envoie le rapport NA par email avec pièce jointe."""
    if not attachment_path:
        print("Aucun fichier à envoyer")
        return

    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = email_recipient
    msg['Subject'] = "Rapport des valeurs manquantes"

    body = "Veuillez trouver ci-joint le rapport des valeurs manquantes."
    msg.attach(MIMEText(body, 'plain'))

    # Ajout de la pièce jointe
    try:
        with open(attachment_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(attachment_path)}"')
        msg.attach(part)
    except Exception as e:
        print(f"Erreur lecture pièce jointe : {e}")
        return

    # Envoi du mail
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(email_sender, email_password)
            server.send_message(msg)
        print(f"Email envoyé avec succès à {email_recipient}")
    except Exception as e:
        print(f"Erreur d'envoi : {e}")
    finally:
        if os.path.exists(attachment_path):
            os.remove(attachment_path)
            print("Fichier temporaire supprimé")

def main():
    """Exécution principale"""
    print("=== Upload et détection des NA ===")
    file_path = upload_excel_file_colab()
    if not file_path:
        return

    na_data = detect_na_in_excel(file_path)
    report_file = create_na_report(na_data)

    if report_file:
        print("=== Envoi du rapport par email ===")
        email_recipient = input("Email destinataire : ")
        email_sender = input("Email expéditeur : ")
        email_password = input("Mot de passe ou token d'application : ")
        send_na_report(email_recipient, report_file, email_sender, email_password)

if __name__ == "__main__":
    main()
