import os
import time
import shutil
import hmac
import pandas as pd
import streamlit as st
from docx import Document

from Replacer import WordReplace


mappings = {
    # '[adresse_client]' : "None",
    # '[objectif_formation]' : "None",
    # '[beneficiaire_formation]' : "None",
    "[telephone_stagiaire]": "Téléphone de contact du stagiaire",
    "[code_postal]": "Code postal du stagiaire",
    "[pays_stagiaire]": "Pays du stagiaire",
    "[email_stagiaire]": "Email du stagiaire",
    "[portable_stagiaire]": "Téléphone de contact du stagiaire",
    "[ville_de_naissance]": "Ville de naissance du stagiaire",
    "[ville_stagiaire]": "Ville du stagiaire",
    "[pays_de_naissance]": "Pays de naissance du stagiaire",
    "[date_de_naissance]": "Date de naissance du stagiaire",
    "[adresse_stagiaire]": "Adresse stagiaire",
    "[Nationalité]": "Nationalité du stagiaire",
    "[nom_stagiaire]": "Nom stagiaire",
    "[fonction_client]": "Fonction du responsable client",
    "[prenom_stagiaire]": "Prénom stagiaire",
    "[nom_organisme] ": "Nom de l'organisme",
    "[nom_responsable]": "Nom du responsable de l'entreprise",
    "[mail]  ": "Email de contact",
    "[ville]": "Ville de l'organisme",
    "[telephone] ": "Téléphone de contact",
    "[region]  ": "Région",
    "[adresse] ": "Adresse complète",
    "[fonction]": "Fonction",
    "[siret] ": "Numéro Siret",
    "[ape] ": "Code APE",
    "[nda] ": "Numéro NDA",
    "[TVA] ": "Numéro TVA",
    "[ville_greffe] ": "Nom du RCS",
    "[site_web]": "Site Internet",
    "[nom-client]": "Responsable du client",
    "[domaine-formation]": "Domaine de formation",
    "[formateur]": "Nom du formateur principal",
    "[nom_referent_handicap] ": "Nom référent handicap",
    "[nom_formation]": "Nom de la formation",
    "[client_entreprise]": "Entreprise cliente bénéficiaire",
    "[nombre_heures]": "Nombre d'heures de la formation",
    "[date_formation]": "Date de la formation",
    "[heures_formation]": "Horaires de la formation",
    "[lieu_formation]": "Lieu de la formation",
    "[prix_formation]": "Prix de la formation",
    "[type_formation]": "Type de la formation",
    "[siret_client]": "Siret du client",
}


def set_date_and_place(doc):
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.text = run.text.replace("[date]", time.strftime("%d/%m/%Y"))
            run.text = run.text.replace("[date_du_jour]", time.strftime("%d/%m/%Y"))
            run.text = run.text.replace("[Fait_a]", "Arles")


def replace_text(doc, old_text, new_text):
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.text = run.text.replace(old_text, new_text)


def check_password():
    """Returns `True` if the user had a correct password."""

    def login_form():
        """Form with widgets to collect user information"""
        with st.form("Credentials"):
            st.text_input("Username", key="username")
            st.text_input("Password", type="password", key="password")
            st.form_submit_button("Log in", on_click=password_entered)

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["username"] in st.secrets[
            "passwords"
        ] and hmac.compare_digest(
            st.session_state["password"],
            st.secrets.passwords[st.session_state["username"]],
        ):
            st.session_state["password_correct"] = True
            # Don't store the username or password.
            del st.session_state["password"]
            del st.session_state["username"]
        else:
            st.session_state["password_correct"] = False

    # Return True if the username + password is validated.
    if st.session_state.get("password_correct", False):
        return True

    # Show inputs for username + password.
    login_form()
    if "password_correct" in st.session_state:
        st.error("😕 User not known or password incorrect")
    return False


if not check_password():
    st.stop()


# Create the Streamlit app
def main():

    if not check_password():
        st.stop()

    st.title("Générateur de dossier Qualiopi :rocket: ")

    st.sidebar.title("Depot d'excel :newspaper: ")

    excel = st.sidebar.file_uploader(
        "Uploader votre fichier excel", type=["csv", "xlsx", "xls"]
    )

    if not excel:
        st.warning("Veuillez uploader un fichier excel pour commencer.")

    if excel:

        df = pd.read_excel(excel)

        st.sidebar.write(df.iloc[:, 1:].tail(7))

        row_index = st.number_input(
            "Entrer l'indice de ligne prévisualisé à gauche",
            min_value=0,
            max_value=len(df) - 1,
            step=1,
        )

        if st.button("Preview du client"):
            row_data = df.iloc[row_index][
                ["Nom de l'organisme", "Nom du responsable de l'entreprise"]
            ]
            st.write(row_data)

        template_folder_path = "templates"

        if st.button("Générer les documents") and template_folder_path:
            # Create a folder to store generated documents
            output_folder_path = (
                f"docs/generated_documents_{time.strftime('%Y%m%d_%H%M%S')}"
            )
            shutil.copytree(
                template_folder_path,
                output_folder_path,
                ignore=shutil.ignore_patterns("*.*"),
            )

            mapping_dict = {
                key: str(df.iloc[row_index][value]) for key, value in mappings.items()
            }

            progress_bar = st.progress(0, text=f"Progress: 0%")
            doc_list = WordReplace.docx_list(template_folder_path)
            for i, file in enumerate(doc_list):
                progress_bar.progress(
                    (i + 1) / len(WordReplace.docx_list(template_folder_path)),
                    text=f"Document numero {i + 1}/{len(doc_list)}",
                )
                wordreplace = WordReplace(file)
                wordreplace.replace_doc(mapping_dict)
                doc = wordreplace.docx
                set_date_and_place(doc)

                nom_prenom = df.iloc[row_index]["Nom de l'organisme"]
                doc_name = f"{nom_prenom}_{os.path.basename(file)}"
                rel_path = os.path.relpath(file, template_folder_path)
                path_to_save = os.path.join(output_folder_path, rel_path)
                path_to_save = path_to_save.replace(os.path.basename(file), doc_name)
                doc.save(path_to_save)

            shutil.make_archive(output_folder_path, "zip", output_folder_path)

            with open(output_folder_path + ".zip", "rb") as f:
                st.download_button(
                    label="Télécharger le dossier des documents générés",
                    data=f,
                    file_name=output_folder_path + ".zip",
                    mime="application/zip",
                )

        if st.button("Supprimer le dossier généré"):
            if os.path.exists("docs"):
                shutil.rmtree("docs")
                print("Folder deleted")


if __name__ == "__main__":
    main()
