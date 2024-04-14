import os
import time
import shutil
import hmac
import pandas as pd
import streamlit as st
from docx import Document

from Replacer import WordReplace


mappings = {
    '[nom_organisme]': 'nom_entreprise',
    '[nom_responsable]': 'nom',
    # '[adresse]': 'adresse_entreprise',
    # '[fonction]': '',
    # '[siret]': '',
    # '[ape]': '',
    # '[nda]': '',
    # '[région]': '',
    # '[ville_greffe]': '',
    # '[telephone]': '',
    # '[mail]': '',
    # '[ville]': '',
    # '[site_web]': '',
    # '[date]': '',
    # '[domaine-formation]': '',
    # '[formateur]': '',
    # '[nom_referent_handicap]': '',
    # '[nom_formation]': '',
    # '[client_entreprise]': '',
    # '[beneficiaire_formation]': '',
    # '[adresse_client]': '',
    # '[nom-client]': '',
    # '[siret_client]': '',
    # '[fonction_client]': '',
    # '[objectif_formation]': '',
    # '[nombre_heures]': '',
    # '[date_formation]': '',
    # '[heures_formation]': '',
    # '[lieu_formation]': '',
    # '[prix_formation]': '',
    # '[type_formation]': '',
    # '[capital]': '',
    # '[nom_associe1]': '',
    # '[adresse_associe1]': '',
    # '[naissance_associe1]': '',
    # '[lieu_naissance_associe1]': '',
    # '[nationalite_associe1]': '',
    # '[situation_associe1]': '',
    # '[apport_associe1]': '',
    # '[nom_associe2]': '',
    # '[adresse_associe2]': '',
    # '[naissance_associe2]': '',
    # '[lieu_naissance_associe2]': '',
    # '[nationalite_associe2]': '',
    # '[situation_associe2]': '',
    # '[apport_associe2]': '',

}


def set_date_and_place(doc):
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.text = run.text.replace("[jour]", time.strftime("%d/%m/%Y"))
            run.text = run.text.replace("[lieu]", 'Arles')


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
        "Uploader votre fichier excel", type=["csv", "xlsx", "xls"])

    if not excel:
        st.warning("Veuillez uploader un fichier excel pour commencer.")

    if excel:

        df = pd.read_excel(excel)

        st.sidebar.write(df.iloc[:, 1:].tail(7))

        columns = df.columns
        new_columns = ['timestamp', 'nom', 'mail', 'tel', 'entreprise_cree', 'nom_entreprise', 'adresse_entreprise',
                       'rqth', 'activite_entreprise', 'capital_social', 'date_bilan', 'associes', 'sturcture_juridique', 'date_debut']
        df.columns = new_columns + list(columns[len(new_columns):])

        row_index = st.number_input(
            "Entrer l'indice de ligne prévisualisé à gauche", min_value=0, max_value=len(df)-1, step=1)

        if st.button("Preview du client"):
            row_data = df.iloc[row_index][["nom", "mail"]]
            st.write(row_data)

        template_folder_path = "templates"

        if st.button('Générer les documents') and template_folder_path:
            # Create a folder to store generated documents
            output_folder_path = f"docs/generated_documents_{time.strftime('%Y%m%d_%H%M%S')}"
            shutil.copytree(template_folder_path, output_folder_path,
                            ignore=shutil.ignore_patterns('*.*'))

            mapping_dict = {key: df.iloc[row_index][value]
                            for key, value in mappings.items()}

            for i, file in enumerate(WordReplace.docx_list(template_folder_path)):
                wordreplace = WordReplace(file)
                wordreplace.replace_doc(mapping_dict)
                doc = wordreplace.docx
                set_date_and_place(doc)

                nom_prenom = df.iloc[row_index]["nom"]
                doc_name = f"{nom_prenom}_{os.path.basename(file)}"
                rel_path = os.path.relpath(file, template_folder_path)
                path_to_save = os.path.join(output_folder_path, rel_path)
                path_to_save = path_to_save.replace(
                    os.path.basename(file), doc_name)
                doc.save(path_to_save)

            shutil.make_archive(output_folder_path, 'zip', output_folder_path)

            with open(output_folder_path+'.zip', "rb") as f:
                st.download_button(
                    label="Télécharger le dossier des documents générés",
                    data=f,
                    file_name=output_folder_path + ".zip",
                    mime="application/zip",
                )

        if st.button("Supprimer le dossier généré"):
            if os.path.exists('docs'):
                shutil.rmtree('docs')
                print("Folder deleted")


if __name__ == "__main__":
    main()