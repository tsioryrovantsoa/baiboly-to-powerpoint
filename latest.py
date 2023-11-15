import os
import json
import argparse

# Configuration des arguments en ligne de commande
parser = argparse.ArgumentParser(description='Trouver les fichiers modifiés dans un répertoire.')
parser.add_argument('directory_path', help='Chemin du répertoire à vérifier')
parser.add_argument('last_modified_db', type=int, help='Dernière date de modification dans la base de données')
args = parser.parse_args()

directory_path = args.directory_path
last_modified_db = args.last_modified_db

modified_files = []

for filename in os.listdir(directory_path):
    file_path = os.path.join(directory_path, filename)
    last_modified = os.path.getmtime(file_path)
    
    if last_modified > last_modified_db:
        modified_files.append(filename)

print(json.dumps(modified_files))
