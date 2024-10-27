import os
import re

# We want to create a function that finds all the folders called "string_1" and replace
# their name with "string_1" + "string_2" where "string_2" is the folder name
# of the grandparent folder of "string_1"


def replace_folder_root_name(root_path, string_1):
    for dirpath, dirnames, filenames in os.walk(root_path):
        for dirname in dirnames:
            if dirname == string_1:
                grandparent_folder = os.path.basename(dirpath)
                grandparent_folder = grandparent_folder.split("_")[:2]
                grandparent_folder = "_".join(grandparent_folder)
                new_dirname = string_1 + "_" + grandparent_folder
                os.rename(
                    os.path.join(dirpath, dirname), os.path.join(dirpath, new_dirname)
                )


replace_folder_root_name("app/templates copy", "PREUVES_Mise_en_oeuvre")
