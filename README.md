# Document Filler - Qualiopi Generator

A Streamlit application for generating Qualiopi certification documents by replacing placeholders in Word and Excel templates with actual client data.

## Features

- **Word Document Processing**: Replaces placeholders in .docx files while preserving formatting
- **Excel Document Processing**: Replaces placeholders in .xlsx and .xls files while preserving formatting
- **Batch Processing**: Processes multiple documents simultaneously
- **Template Management**: Organized template structure with indicators
- **Progress Tracking**: Real-time progress bars for document generation
- **Error Handling**: Graceful handling of corrupted or problematic files
- **ZIP Download**: Automatic packaging of generated documents

## New Feature: Excel Processing

The application now supports processing Excel files in addition to Word documents. Excel files in the template folder will have their placeholders replaced with actual data from the uploaded Excel file, while preserving all formatting, formulas, and structure.

### How it works:

1. **Template Excel Files**: Place Excel files with placeholders in the `app/templates/` directory
2. **Placeholder Format**: Use the same placeholder format as Word documents (e.g., `[NOM_ORGANISME]`, `[DATE]`)
3. **Automatic Processing**: Excel files are automatically processed alongside Word documents
4. **Format Preservation**: All Excel formatting, formulas, charts, and structure are preserved

### Supported Excel Features:

- Text replacement in all cells across all sheets
- Date and place placeholder replacement (`[date]`, `[date_du_jour]`, `[Fait_a]`)
- Preservation of:
  - Cell formatting (colors, fonts, borders)
  - Formulas and calculations
  - Charts and graphs
  - Multiple sheets
  - Cell references and links

## Installation

1. Clone the repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application:
   ```bash
   streamlit run app/app.py
   ```

## Usage

1. Upload an Excel file containing client data
2. Upload a logo (optional)
3. Select the row index for the client data
4. Click "Générer les documents" to process templates
5. Download the generated ZIP file

## Template Structure

```
app/templates/
├── Indicateur_1_Information_du_public/
│   ├── Modeles_de_documents/
│   │   ├── *.docx (Word templates)
│   │   └── *.xlsx (Excel templates)
│   └── PREUVES_Mise_en_oeuvre/
└── ... (32 indicators total)
```

## Dependencies

- streamlit
- pandas
- python-docx
- openpyxl
- tqdm

## Docker Support

Build and run with Docker:

```bash
docker build -t document-filler .
docker run -p 8501:8501 document-filler
```

## Error Handling

The application includes robust error handling:
- Corrupted files are skipped with warnings
- Original files are copied if processing fails
- Detailed error messages for debugging
- Graceful degradation for problematic templates 