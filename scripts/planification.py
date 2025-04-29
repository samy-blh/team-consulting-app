import time
import pandas as pd
import sys
import os
import subprocess
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side
import unicodedata

# Demander fichier et date
fichier = input("Entrez le nom du fichier Excel (ex: liste_techniciens.xlsx) : ")
try:
    df = pd.read_excel(fichier)
except Exception as e:
    print(f"Erreur lors de la lecture du fichier : {e}")
    sys.exit()

date_input = input("Entrez la date de planification (format JJ/MM/AAAA) : ")
try:
    date_cible = datetime.strptime(date_input, "%d/%m/%Y").date()
except ValueError:
    print("Format de date invalide. Utilisez JJ/MM/AAAA.")
    sys.exit()

couleur_base = input("Choisissez une couleur de base pour le dégradé (bleu, vert, rouge) : ").strip().lower()
if couleur_base not in ["bleu", "vert", "rouge"]:
    couleur_base = "bleu"

options = Options()
options.add_argument("--start-maximized")

interventions_planifiees = []

def extraire_interventions(driver, nom, login, onglet_type):
    try:
        driver.find_element(By.LINK_TEXT, onglet_type).click()
        time.sleep(4)

        while True:
            cards = driver.find_elements(By.CLASS_NAME, "intervention")
            if not cards:
                break

            total = len(cards)
            for i in range(total):
                try:
                    cards = driver.find_elements(By.CLASS_NAME, "intervention")
                    card = cards[i]
                    text = card.text
                    lines = text.split("\n")
                    date_line = next((l for l in lines if "Date du RDV" in l), None)
                    if not date_line:
                        continue
                    date_str = date_line.split(":")[1].strip()
                    if len(date_str) == 13:
                        date_str += ":00"
                    rdv_time = datetime.strptime(date_str, "%Y-%m-%d %H:%M")

                    if rdv_time.date() != date_cible:
                        continue

                    card.click()
                    time.sleep(2)

                    statut = "Prévue"
                    debut_intervention = ""
                    jeton_val = ""
                    adresse_client = ""

                    labels = driver.find_elements(By.CLASS_NAME, "label")
                    for label in labels:
                        try:
                            b = label.find_element(By.TAG_NAME, "b")
                            label_title = b.text.strip().lower()
                            texte_complet = label.text.strip()

                            if "début de l'intervention" in label_title:
                                parts = texte_complet.split(":")
                                if len(parts) > 1:
                                    debut_intervention = parts[1].strip()
                                    statut = f"Démarrée à {debut_intervention}"

                            elif "jeton" in label_title:
                                parts = texte_complet.split(":")
                                if len(parts) > 1:
                                    jeton_val = parts[1].strip()

                            elif "adresse" in label_title:
                                try:
                                    adresse_client = label.find_element(By.TAG_NAME, "a").text.strip()
                                except:
                                    adresse_client = texte_complet.split(":")[1].strip()
                        except:
                            continue

                    interventions_planifiees.append({
                        "technicien": nom,
                        "login": login,
                        "jeton": jeton_val,
                        "adresse": adresse_client,
                        "rdv": rdv_time.strftime("%Y-%m-%d %H:%M"),
                        "statut": statut,
                        "heure_actuelle": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "type": onglet_type
                    })

                    driver.back()
                    time.sleep(2)
                except Exception as e:
                    print(f"Erreur sur une intervention : {e}")
                    continue
            break

    except Exception as e:
        print(f"Erreur pour {nom} dans l’onglet {onglet_type} : {e}")

for index, row in df.iterrows():
    nom = row["nom"]
    login = str(row["login"])
    password = str(row["password"])

    print(f"Connexion pour {nom}...")

    driver = webdriver.Chrome(options=options)
    driver.get("https://aboracco.pub.app.ftth.iliad.fr/")
    time.sleep(3)

    inputs = driver.find_elements(By.TAG_NAME, "input")
    inputs[0].send_keys(login)
    inputs[1].send_keys(password)

    time.sleep(1)
    bouton_connexion = driver.find_element(By.XPATH, "//button[contains(text(), 'Connexion')]")
    bouton_connexion.click()
    time.sleep(4)

    extraire_interventions(driver, nom, login, "Production")
    extraire_interventions(driver, nom, login, "Post-Production / SAV")

    driver.quit()

if interventions_planifiees:
    date_nom_fichier = date_cible.strftime("%d-%m")
    nom_fichier = f"planning pour le {date_nom_fichier}.xlsx"

    df_result = pd.DataFrame(interventions_planifiees)
    df_result.to_excel(nom_fichier, index=False)

    wb = load_workbook(nom_fichier)
    ws = wb.active

    # Ajustement automatique des colonnes
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2

    # Dégradé par technicien
    couleurs = {
        "bleu": ["CCE5FF", "99CCFF", "66B2FF", "3399FF"],
        "vert": ["CCFFCC", "99FF99", "66FF66", "33FF33"],
        "rouge": ["FFCCCC", "FF9999", "FF6666", "FF3333"]
    }

    couleur_steps = couleurs.get(couleur_base, couleurs["bleu"])
    techniciens = list(set([row[0].value for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1)]))
    technicien_couleurs = {tech: couleur_steps[i % len(couleur_steps)] for i, tech in enumerate(techniciens)}

    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        tech_name = row[0].value
        fill_color = PatternFill(start_color=technicien_couleurs.get(tech_name, "FFFFFF"), end_color=technicien_couleurs.get(tech_name, "FFFFFF"), fill_type="solid")
        for cell in row:
            cell.border = border
            cell.fill = fill_color

    wb.save(nom_fichier)
    print(f"\nFichier '{nom_fichier}' généré avec succès.")

    # Ouvrir automatiquement le fichier
    subprocess.Popen(["start", "", nom_fichier], shell=True)
else:
    print("\nAucune intervention planifiée pour la date spécifiée.")
