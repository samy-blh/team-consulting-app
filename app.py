import streamlit as st
import subprocess
import os
from pathlib import Path
from datetime import datetime

st.set_page_config(page_title="Team Consulting App", layout="wide")

st.title("📋 Outil de gestion - Team Consulting")

# 📂 Liste des fichiers dans 'data/'
data_dir = Path("data")
scripts_dir = Path("scripts")
output_dir = Path("output")

liste_fichiers = sorted([f.name for f in data_dir.glob("*.xlsx")])

if not liste_fichiers:
    st.error("Aucune liste de techniciens trouvée dans 'data/'.")
    st.stop()

# ✅ Sélection de la liste de techniciens
liste_choisie = st.selectbox("🧾 Sélectionnez votre liste de techniciens :", liste_fichiers)

# ✅ Choix de l'action
action = st.selectbox("🔧 Quelle action souhaitez-vous réaliser ?", ["planification", "verification", "terminees"])

# 📅 Date du jour pour les scripts qui en ont besoin
date_du_jour = datetime.now().strftime("%d/%m/%Y")

# Bouton de validation
if st.button("Lancer le traitement"):
    with st.spinner('⏳ Génération du fichier, veuillez patienter...'):

        fichier_liste = data_dir / liste_choisie
        nom_liste = Path(liste_choisie).stem
        dossier_output = output_dir / nom_liste
        dossier_output.mkdir(parents=True, exist_ok=True)

        fichier_sortie = dossier_output / f"{action}.xlsx"

        # Commande à exécuter
        cmd = ["python", str(scripts_dir / f"{action}.py"), str(fichier_liste), str(fichier_sortie)]

        # Ajouter la date uniquement pour planification et terminees
        if action in ["planification", "terminees"]:
            cmd.append(date_du_jour)

        # Exécution
        try:
            subprocess.run(cmd, check=True)
            st.success(f"✅ Fichier généré avec succès : {fichier_sortie.name}")

            with open(fichier_sortie, "rb") as f:
                st.download_button(
                    label="📥 Télécharger le fichier Excel",
                    data=f,
                    file_name=fichier_sortie.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except subprocess.CalledProcessError as e:
            st.error(f"❌ Une erreur est survenue lors du traitement : {e}")

