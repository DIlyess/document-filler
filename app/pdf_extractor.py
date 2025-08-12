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
            "convention": "Numéro de convention",
            "date_signature": "Date de signature",
            "date_debut": "Date de début",
            "date_fin": "Date de fin",
            "organisme": "Nom de l'organisme",
            "responsable": "Responsable de l'organisme",
            "adresse": "Adresse de l'organisme",
            "telephone": "Téléphone",
            "email": "Email",
            "montant": "Montant de la convention",
            "devise": "Devise",
            "objectif": "Objectif de la convention",
            "partenaire": "Partenaire",
            "signataire": "Signataire",
            "reference": "Référence",
            "type": "Type de convention",
            "statut": "Statut de la convention"
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
            value = self.find_text_after_keyword(text, keyword)
            extracted_data[description] = value if value else "Non trouvé"

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

    def export_to_json(self, extracted_data: Dict[str, str]) -> str:
        """
        Convert extracted data to JSON format.

        Args:
            extracted_data: Dictionary with extracted data

        Returns:
            str: JSON formatted string
        """
        import json
        return json.dumps(extracted_data, ensure_ascii=False, indent=2)


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
