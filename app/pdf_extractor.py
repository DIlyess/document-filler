import PyPDF2
import re
from typing import Dict
import streamlit as st


class PDFExtractor:
    """
    Utility class for extracting information from PDF documents based on keywords.
    """

    def __init__(self):
        # Dictionary mapping keywords to field descriptions
        self.extraction_fields = {
            "L'organisme de formation": "Nom de l'organisme",
            "Représentée par": "Prénom et Nom du responsable de l'organisme",
            "Email": "Email de votre organisme de formation",
            "Tel : ": "Téléphone de contact",
            "Siège social au : ": "Ville de l'organisme",
            "": "Région",
            "Siège social au : ": "Adresse (N° + Nom de la rue + code postal + ville)",
            "Représentée par ": "Fonction du dirigeant de l'organisme",
            "Siret : ": "Numéro Siret",
            "": "Code APE",
            "": "Numéro NDA Si vous n'avez pas encore votre NDA mettre la mention  NDA en cours d'enregistrement",
            "- TVA : ": "Numéro TVA Si vous n'êtes pas soumis à TVA mettre xxx",
            "- RCS": "Ville du RCS (Registre du commerce dont dépend votre structure)",
            "": "Site Internet (Mettre le lien direct)",
            "- Intitulé de l’action :": "Nom de la formation",
            "": "Domaine de formation",
            "- Formateur :": "Nom du formateur principal",
            "": "Nom référent handicap (le dirigeant est souvent le réfèrent handicap)",
            "Dates et horaires : ": "Horaires de la formation",
            "Durée de l’action de formation :": "Nombre d'heures de la formation (Indiquez que le chiffre)",
            "Lieu : ": "Lieu de la formation",
            "Dates et horaires : ": "Date début de la formation",
            "": "Nombre de participants",
            "TOTAL GENERAL :": "Prix de la formation",
            "2)": "Entreprise cliente bénéficiaire",
            "": "Responsable du client",
            "": "Siret du client",
            "": "Fonction du responsable client",
            "": "Nom stagiaire",
            "": "Prénom stagiaire",
            "": "Adresse stagiaire",
            "": "Date de naissance du stagiaire",
            "": "Ville de naissance du stagiaire",
            "": "Pays de naissance du stagiaire",
            "": "Nationalité du stagiaire",
            "": "Code postal du candidat/entreprise",
            "": "Ville du stagiaire",
            "": "Pays du stagiaire",
            "": "Téléphone de contact du stagiaire",
            "": "Email du stagiaire",
            "": "Date actualisation de vos documents  (Indiquez mois/année)",
            "": "Code postal organisme de formation",
            "": "Facebook  (Mettre le lien direct)",
            "": "LinkedIn  (Mettre le lien direct)",
            "": "Twitter of",
            "": "Instagram  (Mettre le lien direct)",
            "": "Qualification du formateur",
            "": "Nombre de jours de la formation",
            "": "Date de la signature du contrat/convention",
            "": "Public visé",
            "": "Statut juridique de votre organisme de formation",
            "": "Année",
            "": "[NOM(S) DU/des STAGIAIRE(s)]",
            "": "Effectif stagiaires",
            "": "Date de la fin de la formation",
        }

    def extract_text_from_pdf(self, pdf_file) -> str:
        """
        Extract all text from a PDF file.

        Args:
            pdf_file: Uploaded PDF file from Streamlit

        Returns:
            str: Extracted text from the PDF
        """
        try:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            text = ""

            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text += page.extract_text() + "\n"

            return text
        except Exception as e:
            st.error(f"Erreur lors de la lecture du PDF: {str(e)}")
            return ""

    def find_text_after_keyword(self, text: str, keyword: str, max_chars: int = 100) -> str:
        """
        Find text that appears after a specific keyword.

        Args:
            text: Full text to search in
            keyword: Keyword to search for
            max_chars: Maximum number of characters to extract after the keyword

        Returns:
            str: Text found after the keyword
        """
        # Create a case-insensitive pattern
        pattern = re.compile(re.escape(keyword), re.IGNORECASE)
        match = pattern.search(text)

        if match:
            start_pos = match.end()
            end_pos = min(start_pos + max_chars, len(text))
            extracted_text = text[start_pos:end_pos].strip()

            # Clean up the extracted text (remove extra whitespace, newlines)
            extracted_text = re.sub(r'\s+', ' ', extracted_text)

            # Try to find a natural break point (period, comma, newline)
            # Prioritize newline and period as primary break points
            break_chars = ['\n', '.', ',', ';', ':', ' ']
            for char in break_chars:
                pos = extracted_text.find(char)
                if pos != -1 and pos < 100:  # Allow slightly more characters for better context
                    extracted_text = extracted_text[:pos].strip()
                    break

            return extracted_text

        return ""

    def extract_all_fields(self, text: str) -> Dict[str, str]:
        """
        Extract all fields based on the predefined keywords.

        Args:
            text: Full text from the PDF

        Returns:
            Dict[str, str]: Dictionary with field descriptions as keys and extracted values as values
        """
        extracted_data = {}

        for keyword, description in self.extraction_fields.items():
            if keyword != "":
                value = self.find_text_after_keyword(text, keyword)
                extracted_data[description] = value if value else "Non trouvé"
            else:
                extracted_data[description] = "Non défini"

        return extracted_data

    def export_to_csv(self, extracted_data: Dict[str, str]) -> str:
        """
        Convert extracted data to CSV format.

        Args:
            extracted_data: Dictionary with extracted data

        Returns:
            str: CSV formatted string
        """
        csv_lines = ["Champ,Valeur"]
        for field, value in extracted_data.items():
            # Escape commas and quotes in the value
            escaped_value = value.replace('"', '""')
            csv_lines.append(f'"{field}","{escaped_value}"')

        return "\n".join(csv_lines)


def validate_pdf_file(uploaded_file) -> bool:
    """
    Validate that the uploaded file is a valid PDF.

    Args:
        uploaded_file: File uploaded through Streamlit

    Returns:
        bool: True if valid PDF, False otherwise
    """
    if uploaded_file is None:
        return False

    # Check file extension
    if not uploaded_file.name.lower().endswith('.pdf'):
        return False

    # Check file size (max 50MB)
    if uploaded_file.size > 50 * 1024 * 1024:
        return False

    return True
