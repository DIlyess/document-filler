import os
import re
import unidecode


def sanitize_name(name):
    # Remove accents
    name = unidecode.unidecode(name)
    # Split the name into base and extension
    if "." in name:
        base, ext = name.rsplit(".", 1)
        base = re.sub(r"[^a-zA-Z0-9]", "_", base)
        base = re.sub(r"_+", "_", base)
        return f"{base}.{ext}"
    else:
        name = re.sub(r"[^a-zA-Z0-9]", "_", name)
        name = re.sub(r"_+", "_", name)
        return name


def sanitize_folder_name(name):
    # Remove accents
    name = unidecode.unidecode(name)
    # Replace all non-alphanumeric characters with '_'
    name = re.sub(r"[^a-zA-Z0-9]", "_", name)
    # Remove duplicate underscores
    name = re.sub(r"_+", "_", name)
    return name


def strip_spaces_and_sanitize_recursively(root_path):
    for dirpath, dirnames, filenames in os.walk(root_path, topdown=False):
        # Process files
        for filename in filenames:
            new_filename = sanitize_name(filename)
            if new_filename != filename:
                os.rename(
                    os.path.join(dirpath, filename), os.path.join(dirpath, new_filename)
                )

        # Process directories
        for dirname in dirnames:
            new_dirname = sanitize_folder_name(dirname)
            if new_dirname != dirname:
                os.rename(
                    os.path.join(dirpath, dirname), os.path.join(dirpath, new_dirname)
                )


strip_spaces_and_sanitize_recursively("templates")
