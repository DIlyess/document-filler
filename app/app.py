import os
import streamlit as st
from tqdm import tqdm
import concurrent.futures
import pandas as pd
import time
import shutil
import gc

from Replacer import WordReplace
from ExcelReplacer import ExcelReplace
from pdf_extractor import PDFExtractor, validate_pdf_file
from utils import (
    zip_folder, create_mapping_dict, set_date_and_place,
    replace_text, replace_first_image_in_header
)


def process_word_document(args):
    """
    Process a single Word document - designed for parallel execution
    """
    file_path, mapping_dict, logo_path, template_folder_path, output_folder_path = args

    try:
        wordreplace = WordReplace(file_path)
        wordreplace.replace_doc(mapping_dict)
        doc = wordreplace.docx
        set_date_and_place(doc)

        if logo_path and os.path.exists(logo_path):
            replace_first_image_in_header(doc, logo_path)

        doc_name = f"{os.path.basename(file_path)}"
        rel_path = os.path.relpath(file_path, template_folder_path)
        path_to_save = os.path.join(output_folder_path, rel_path)
        path_to_save = path_to_save.replace(
            os.path.basename(file_path), doc_name)
        doc.save(path_to_save)
        return True, file_path
    except Exception as e:
        return False, f"Error processing {os.path.basename(file_path)}: {str(e)}"


def process_excel_document(args):
    """
    Process a single Excel document - designed for parallel execution
    """
    file_path, mapping_dict, template_folder_path, output_folder_path = args

    try:
        excel_replace = ExcelReplace(file_path)
        excel_replace.replace_excel(mapping_dict)
        excel_replace.set_date_and_place()

        excel_name = f"{os.path.basename(file_path)}"
        rel_path = os.path.relpath(file_path, template_folder_path)
        path_to_save = os.path.join(output_folder_path, rel_path)
        path_to_save = path_to_save.replace(
            os.path.basename(file_path), excel_name)
        excel_replace.save(path_to_save)
        return True, file_path
    except Exception as e:
        return False, f"Error processing {os.path.basename(file_path)}: {str(e)}"


# Create the Streamlit app
def main():

    # if not check_password():
    #     st.stop()

    st.title("Générateur de dossier Qualiopi :rocket: ")

    # Create tabs for different functionalities
    tab1, tab2 = st.tabs(
        ["Génération de documents", "Extraire les informations de la convention"])

    with tab1:
        st.header("Génération de documents")

        st.sidebar.title("Depot d'excel :newspaper: ")

        excel = st.sidebar.file_uploader(
            "Uploader votre fichier excel", type=["csv", "xlsx", "xls"]
        )

        logo = st.sidebar.file_uploader(
            "Uploader votre logo", type=["png", "jpg", "jpeg"])

        # Performance configuration
        st.sidebar.title("Configuration Performance :gear:")
        use_parallel = st.sidebar.checkbox("Utiliser le traitement parallèle", value=True,
                                           help="Active le traitement parallèle pour améliorer les performances")
        max_workers = st.sidebar.slider("Nombre de workers parallèles", min_value=1,
                                        max_value=8, value=4, help="Nombre de documents traités simultanément")

        if not excel:
            st.warning("Veuillez uploader un fichier excel pour commencer.")

        if excel:

            if logo is not None:
                # Save the uploaded file to the current directory
                with open("logo.png", "wb") as f:
                    f.write(logo.getvalue())

            df = pd.read_excel(excel)
            df = df.astype(str)

            st.sidebar.write(df.iloc[:, 1:].tail(7))

            row_index = st.number_input(
                "Entrer l'indice de ligne prévisualisé à gauche",
                min_value=0,
                max_value=len(df) - 1,
                step=1,
            )

            if st.button("Preview du client"):
                row_data = df.iloc[row_index][
                    ["Nom de l'organisme", "Prénom et Nom du responsable de l'organisme"]
                ]
                st.write(row_data)

            if os.path.exists("templates"):
                template_folder_path = "templates"
            else:
                template_folder_path = "app/templates"

            if st.button("Générer les documents") and template_folder_path:
                start_time = time.time()

                nom_organisme = df.iloc[row_index]["Nom de l'organisme"]
                # Create a folder to store generated documents
                output_folder_path = f"docs/{nom_organisme}_{time.strftime('%H_%M_%S')}"
                # Copy template structure (now including Excel files)
                shutil.copytree(
                    template_folder_path,
                    output_folder_path,
                    ignore=shutil.ignore_patterns("*.docx"),
                )

                mappings = create_mapping_dict(df)

                mapping_dict = {
                    key: str(df.iloc[row_index][value]) for key, value in mappings.items()
                }

                progress_bar = st.progress(0, text=f"Progress: 0%")

                # Process both Word and Excel documents
                doc_list = WordReplace.docx_list(template_folder_path)
                excel_list = ExcelReplace.excel_list(template_folder_path)
                total_files = len(doc_list) + len(excel_list)

                st.info(
                    f"Traitement de {len(doc_list)} documents Word et {len(excel_list)} documents Excel")

                file_counter = 0

                # Process Word documents
                if use_parallel and len(doc_list) > 1:
                    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                        futures = [
                            executor.submit(process_word_document, (file, mapping_dict,
                                            "logo.png" if logo is not None else None, template_folder_path, output_folder_path))
                            for file in doc_list
                        ]
                        for future in tqdm(concurrent.futures.as_completed(futures), total=len(futures)):
                            file_counter += 1
                            progress_bar.progress(
                                file_counter / total_files,
                                text=f"Document Word {file_counter}/{total_files}",
                            )
                            success, result = future.result()
                            if not success:
                                # result contains the error message
                                st.warning(result)
                            # Force garbage collection to free memory
                            gc.collect()
                else:
                    # Sequential processing for small document sets or when parallel is disabled
                    for i, file in tqdm(enumerate(doc_list)):
                        file_counter += 1
                        progress_bar.progress(
                            file_counter / total_files,
                            text=f"Document Word {file_counter}/{total_files}",
                        )
                        success, result = process_word_document(
                            (file, mapping_dict, "logo.png" if logo is not None else None, template_folder_path, output_folder_path))
                        if not success:
                            st.warning(result)
                        gc.collect()

                # Process Excel documents (new feature)
                if use_parallel and len(excel_list) > 1:
                    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                        futures = [
                            executor.submit(
                                process_excel_document, (file, mapping_dict, template_folder_path, output_folder_path))
                            for file in excel_list
                        ]
                        for future in tqdm(concurrent.futures.as_completed(futures), total=len(futures)):
                            file_counter += 1
                            progress_bar.progress(
                                file_counter / total_files,
                                text=f"Document Excel {file_counter}/{total_files}",
                            )
                            success, result = future.result()
                            if not success:
                                # result contains the error message
                                st.warning(result)
                            # Force garbage collection to free memory
                            gc.collect()
                else:
                    # Sequential processing for small document sets or when parallel is disabled
                    for i, file in tqdm(enumerate(excel_list)):
                        file_counter += 1
                        progress_bar.progress(
                            file_counter / total_files,
                            text=f"Document Excel {file_counter}/{total_files}",
                        )
                        success, result = process_excel_document(
                            (file, mapping_dict, template_folder_path, output_folder_path))
                        if not success:
                            st.warning(result)
                        gc.collect()

                zip_folder(output_folder_path, output_folder_path + ".zip")

                end_time = time.time()
                processing_time = end_time - start_time

                st.success(
                    f"Documents générés en {processing_time:.2f} secondes ! Dossier: {output_folder_path}")
                st.info(
                    f"Performance: {total_files/processing_time:.2f} documents/seconde")

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
                    # print("Folder deleted")

    # PDF Information Extraction Tab
    with tab2:
        st.header("Extraire les informations de la convention")
        st.write(
            "Téléchargez un fichier PDF de convention pour extraire automatiquement les informations importantes.")

        # File upload
        uploaded_pdf = st.file_uploader(
            "Choisir un fichier PDF",
            type=['pdf'],
            help="Sélectionnez un fichier PDF de convention (max 50MB)"
        )

        if uploaded_pdf is not None:
            # Validate PDF file
            if not validate_pdf_file(uploaded_pdf):
                st.error("Veuillez sélectionner un fichier PDF valide (max 50MB)")
            else:
                st.success(f"Fichier PDF chargé : {uploaded_pdf.name}")

                # Initialize PDF extractor
                extractor = PDFExtractor()

                # Extract text from PDF
                with st.spinner("Extraction du texte du PDF en cours..."):
                    extracted_text = extractor.extract_text_from_pdf(
                        uploaded_pdf)

                if extracted_text:
                    # Extract all default fields automatically
                    extracted_data = extractor.extract_all_fields(
                        extracted_text)

                    # Create a horizontal table
                    st.subheader("Informations extraites de la convention")

                    # Create a DataFrame for better table display and transpose it for horizontal view
                    df_results = pd.DataFrame(list(extracted_data.items()), columns=[
                                              'Champ', 'Valeur']).set_index('Champ')

                    # Set the same order as extractor.extraction_fields
                    col_in_common = list(
                        set(extractor.extraction_fields.values()) & set(df_results.index))

                    df_results = df_results.loc[col_in_common].T

                    # Display the transposed table
                    st.dataframe(
                        df_results, use_container_width=True, hide_index=True)

                    # Export options
                    st.subheader("Exporter les résultats")
                    col1, col2 = st.columns(2)

                    with col1:
                        csv_data = extractor.export_to_csv(extracted_data)
                        st.download_button(
                            label="📥 Télécharger en CSV",
                            data=csv_data,
                            file_name="convention_extracted_data.csv",
                            mime="text/csv"
                        )
                else:
                    st.error(
                        "Impossible d'extraire le texte du PDF. Vérifiez que le fichier n'est pas corrompu.")


if __name__ == "__main__":
    main()
