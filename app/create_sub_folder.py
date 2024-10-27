# We want to look for all folder that contains the patttern "Indicateur" in the name
# and create a subfolder in this folder called "Preuves_Mise_en_Oeuvre"+parent_folder_name

import os
import shutil


def create_sub_folder(parent_folder_path):
    for root, dirs, files in os.walk(parent_folder_path):
        for dir in dirs:
            if "Indicateur_" in dir and "Preuves_Mise_en_Oeuvre" not in dir:
                parent_name = "_".join(dir.split("_")[:2])
                sub_folder_name = "Preuves_Mise_en_Oeuvre" + "_" + parent_name
                new_folder_path = os.path.join(root, dir, sub_folder_name)
                os.makedirs(new_folder_path, exist_ok=True)
                print(f"Created folder: {new_folder_path}")


create_sub_folder("app/templates")

# Now we want to add a file "Consigne.txt" to each of the newly created subfolders


def add_consigne_txt(parent_folder_path):
    for root, dirs, files in os.walk(parent_folder_path):
        for dir in dirs:
            if "Preuves_Mise_en_Oeuvre_Indicateur" in dir:
                consigne_file_path = os.path.join(root, dir, "Consigne.txt")
                with open(consigne_file_path, "w") as f:
                    f.write("Veuillez ajouter ici les preuves de mise en oeuvre")


add_consigne_txt("app/templates")
