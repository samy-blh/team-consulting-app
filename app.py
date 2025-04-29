import streamlit as st
import os
from pathlib import Path

# Config Streamlit
st.set_page_config(page_title="Team Consulting App", layout="wide")

st.title("📊 Interface - Team Consulting & Co")

# Dossiers
liste_dossiers = sorted([f.name for f in Path("output").iterdir() if f.is_dir()])

if not liste_dossiers:
    st.warning("Aucune liste de techniciens trouvée dans le dossier 'output/'.")
    st.stop()

# Choix de la liste
liste_choisie = st.selectbox("🧾 Choisissez une liste de techniciens :", liste_dossiers)

# Choix de l'action
action = st.selectbox("🔧 Quelle action souhaitez-vous consulter ?", ["planification", "verification", "terminees"])

# Déduire le chemin du fichier
fichier_excel = f"output/{liste_choisie}/{action}.xlsx"

if os.path.exists(fichier_excel):
    with open(fichier_excel, "rb") as f:
        st.download_button(
            label="📥 Télécharger le fichier Excel",
            data=f,
            file_name=f"{action}_{liste_choisie}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.error(f"Aucun fichier disponible pour : `{action}` dans `{liste_choisie}`.")
