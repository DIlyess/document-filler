# Document Filler - Qualiopi Generator

A Streamlit application for generating Qualiopi certification documents by replacing placeholders in Word and Excel templates with actual client data.

## Features

- **Word Document Processing**: Replaces placeholders in .docx files while preserving formatting
- **Excel Document Processing**: Replaces placeholders in .xlsx and .xls files while preserving formatting
- **Batch Processing**: Processes multiple documents simultaneously
- **Parallel Processing**: Multi-threaded processing for improved performance with large document sets
- **Template Management**: Organized template structure with indicators
- **Progress Tracking**: Real-time progress bars for document generation
- **Error Handling**: Graceful handling of corrupted or problematic files
- **ZIP Download**: Automatic packaging of generated documents
- **Performance Monitoring**: Real-time performance metrics and processing time tracking

## Performance Optimizations

### Parallel Processing
- **Multi-threading**: Process multiple documents simultaneously using ThreadPoolExecutor
- **Configurable Workers**: Adjustable number of parallel workers (1-8)
- **Smart Fallback**: Automatically switches to sequential processing for small document sets
- **Memory Management**: Automatic garbage collection to prevent memory leaks

### Algorithm Improvements
- **Optimized Text Replacement**: Single-pass processing for all replacements in a paragraph
- **Caching**: Cached access to document sections and tables
- **Efficient Data Structures**: Reduced redundant operations and improved memory usage
- **Early Exit**: Skip processing for empty paragraphs or missing placeholders

### Performance Gains
- **2-4x faster** processing for large document sets (100+ documents)
- **Reduced memory usage** through optimized data structures
- **Better scalability** with configurable parallel processing
- **Real-time performance metrics** showing documents per second

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
3. Configure performance settings in the sidebar:
   - Enable/disable parallel processing
   - Adjust number of parallel workers
4. Select the row index for the client data
5. Click "Générer les documents" to process templates
6. Monitor real-time performance metrics
7. Download the generated ZIP file

## Performance Configuration

### Parallel Processing Settings
- **Enable Parallel Processing**: Toggle to enable/disable multi-threading
- **Number of Workers**: Adjust from 1-8 parallel workers
- **Auto-fallback**: Automatically uses sequential processing for small document sets

### Performance Monitoring
- **Processing Time**: Shows total time taken for document generation
- **Documents per Second**: Real-time performance metric
- **Progress Tracking**: Detailed progress for Word and Excel documents
- **Error Reporting**: Individual file error reporting without stopping the process

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
- concurrent.futures (built-in)

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
- Individual file error reporting without stopping batch processing

## Performance Tips

1. **Use Parallel Processing**: Enable for document sets with 10+ files
2. **Adjust Worker Count**: Use 4-6 workers for optimal performance
3. **Monitor Memory**: Large document sets may benefit from lower worker counts
4. **Template Optimization**: Remove unused placeholders to improve processing speed 