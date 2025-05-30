import time
import sys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side
from pathlib import Path

# Arguments attendus : fichier technicien, fichier sortie, date cible
fichier_excel = sys.argv[1]
fichier_sortie = sys.argv[2]
date_input = sys.argv[3]

date_cible = datetime.strptime(date_input, "%d/%m/%Y").date()
couleur_base = "bleu"  # couleur par défaut

options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

df = pd.read_excel(fichier_excel)

interventions_planifiees = []

def extraire_interventions(driver, nom, login, onglet_type):
    try:
        driver.find_element(By.LINK_TEXT, onglet_type).click()
        time.sleep(4)
        cards = driver.find_elements(By.CLASS_NAME, "intervention")

        for card in cards:
            try:
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

                debut_intervention = ""
                jeton_val = ""
                adresse_client = ""

                labels = driver.find_elements(By.CLASS_NAME, "label")
                for label in labels:
                    try:
                        b = label.find_element(By.TAG_NAME, "b")
                        label_title = b.text.strip().lower()
                        texte_complet = label.text.strip()

                        if "début" in label_title:
                            debut_intervention = texte_complet.split(":")[1].strip()
                        elif "jeton" in label_title:
                            jeton_val = texte_complet.split(":")[1].strip()
                        elif "adresse" in label_title:
                            adresse_client = texte_complet.split(":")[1].strip()
                    except:
                        continue

                interventions_planifiees.append({
                    "technicien": nom,
                    "login": login,
                    "jeton": jeton_val,
                    "adresse": adresse_client,
                    "rdv": rdv_time.strftime("%Y-%m-%d %H:%M"),
                    "statut": "Prévue",
                    "heure_actuelle": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "type": onglet_type
                })

                driver.back()
                time.sleep(2)

            except:
                continue

    except:
        pass

for index, row in df.iterrows():
    nom = row["nom"]
    login = str(row["login"])
    password = str(row["password"])

    driver = webdriver.Chrome(options=options)
    driver.get("https://aboracco.pub.app.ftth.iliad.fr/")
    time.sleep(3)

    inputs = driver.find_elements(By.TAG_NAME, "input")
    inputs[0].send_keys(login)
    inputs[1].send_keys(password)
    driver.find_element(By.XPATH, "//button[contains(text(), 'Connexion')]").click()
    time.sleep(4)

    extraire_interventions(driver, nom, login, "Production")
    extraire_interventions(driver, nom, login, "Post-Production / SAV")
    driver.quit()

if interventions_planifiees:
    Path(fichier_sortie).parent.mkdir(parents=True, exist_ok=True)
    df_result = pd.DataFrame(interventions_planifiees)
    df_result.to_excel(fichier_sortie, index=False)
