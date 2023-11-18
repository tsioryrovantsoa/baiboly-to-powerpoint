import sys
import os
import subprocess

# Récupérer le chemin du fichier PowerPoint à partir des arguments de la ligne de commande
if len(sys.argv) < 2:
    print("Veuillez fournir le chemin du fichier PowerPoint en tant qu'argument.")
    sys.exit(1)

# Utiliser os.path.abspath pour normaliser le chemin d'accès
file_path = os.path.abspath(" ".join(sys.argv[1:]))

# Lancer le fichier PowerPoint en arrière-plan sans bloquer le script Python
subprocess.Popen(["start", "", file_path], shell=True)
