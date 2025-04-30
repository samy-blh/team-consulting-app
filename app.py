import streamlit as st
import subprocess
import time
from pathlib import Path
from datetime import datetime

st.set_page_config(page_title="Team Consulting App", layout="wide")

st.title("📋 Outil de gestion - Team Consulting")

# 📁 Dossiers
data_dir = Path("data")
scripts_dir = Path("scripts")
output_dir = Path("output")

liste_fichiers = sorted([f.name for f in data_dir.glob("*.xlsx")])

if not liste_fichiers:
    st.error("Aucune liste de techniciens trouvée dans 'data/'.")
    st.stop()

# ✅ Sélection de la liste
liste_choisie = st.selectbox("🧾 Sélectionnez votre liste de techniciens :", liste_fichiers)

# ✅ Choix de l'action
action = st.selectbox("🔧 Quelle action souhaitez-vous réaliser ?", ["planification", "verification", "terminees"])

# 📅 Date du jour pour les scripts
date_du_jour = datetime.now().strftime("%d/%m/%Y")

# ✅ Lancer le traitement
if st.button("Lancer le traitement"):
    with st.spinner("⏳ Traitement en cours..."):

        fichier_liste = data_dir / liste_choisie
        nom_liste = Path(liste_choisie).stem
        dossier_output = output_dir / nom_liste
        dossier_output.mkdir(parents=True, exist_ok=True)

        fichier_sortie = dossier_output / f"{action}.xlsx"

        cmd = ["python", str(scripts_dir / f"{action}.py"), str(fichier_liste), str(fichier_sortie)]
        if action in ["planification", "terminees"]:
            cmd.append(date_du_jour)

        st.code(f"Commande exécutée : {' '.join(cmd)}")

        try:
            result = subprocess.run(cmd, check=True, capture_output=True, text=True)

            # Pause pour éviter FileNotFoundError
            time.sleep(2)

            st.success(f"✅ Fichier généré avec succès : {fichier_sortie.name}")
            with open(fichier_sortie, "rb") as f:
                st.download_button(
                    label="📥 Télécharger le fichier Excel",
                    data=f,
                    file_name=fichier_sortie.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.code(f"Sortie standard : {result.stdout}")
            st.code(f"Erreur standard : {result.stderr}")

        except subprocess.CalledProcessError as e:
            st.error("❌ Une erreur est survenue.")
            st.code(f"Commande exécutée : {' '.join(cmd)}")
            st.code(f"Code de sortie : {e.returncode}")
            st.code(f"Erreur : {e.stderr}")
