import subprocess
from pathlib import Path
from datetime import datetime

data_dir = Path("data")
output_dir = Path("output")
scripts_dir = Path("scripts")

date_du_jour = datetime.now().strftime("%d/%m/%Y")

actions = {
    "planification": "planification.py",
    "verification": "verification.py",
    "terminees": "terminees.py"
}

liste_fichiers = sorted([f for f in data_dir.glob("*.xlsx")])

for fichier in liste_fichiers:
    nom_liste = fichier.stem
    print(f"\n📂 Traitement de la liste : {nom_liste}")

    dossier_resultat = output_dir / nom_liste
    dossier_resultat.mkdir(parents=True, exist_ok=True)

    for action, script_name in actions.items():
        chemin_script = scripts_dir / script_name
        chemin_sortie = dossier_resultat / f"{action}.xlsx"

        print(f"🔄 Lancement de : {action.upper()} pour {nom_liste}")
        print(f"📁 Script : {chemin_script}")
        print(f"📄 Entrée : {fichier}")
        print(f"📤 Sortie : {chemin_sortie}")

        try:
            cmd = ["python", str(chemin_script), str(fichier), str(chemin_sortie)]
            if action in ["planification", "terminees"]:
                cmd.append(date_du_jour)

            subprocess.run(cmd, check=True)
            print(f"✅ {action} - {nom_liste} OK")

        except subprocess.CalledProcessError as e:
            print(f"❌ Erreur dans {action} - {nom_liste} : {e}")
