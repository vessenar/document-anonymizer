# Document Anonymizer Tool

The **Document Anonymizer Tool** is a Python-based script designed to help freelancers, solopreneurs, and small businesses anonymize sensitive information (PII - Personally Identifiable Information) in their portfolio documents. This tool automates the process of redacting sensitive data, saving time and effort when preparing documents for sharing with clients or the public.

## Problem Statement

Manually anonymizing documents is a tedious and error-prone task, especially for freelancers and solopreneurs who need to showcase their work while protecting sensitive information. This tool was created to address this challenge, particularly for those in **B2B marketing**, where sharing anonymized case studies, proposals, and reports is common.

## Features

- **Batch Processing**: Anonymize multiple files in a folder at once.
- **Supported File Types**:
  - Microsoft Word (`.docx`)
  - Microsoft Excel (`.xlsx`)
  - PDF files
- **Customizable Redaction**:
  - Automatically detects and redacts entities such as names, organizations, locations, and products using [spaCy](https://spacy.io/).
  - Allows users to define a **whitelist** of keywords that should not be redacted.
- **Redaction Log**: Generates a CSV log of all redacted entities for transparency and review.

## How It Works

1. The script uses the **spaCy NLP library** to identify entities such as:
   - `PERSON` (e.g., names)
   - `ORG` (e.g., organizations)
   - `GPE` (e.g., locations)
   - `PRODUCT` (e.g., product names)
2. Detected entities are replaced with placeholders (e.g., `[PERSON]`, `[ORG_NAME]`).
3. The tool processes:
   - **Paragraphs** in Word documents
   - **Cells** in Excel spreadsheets
   - **Text** in PDFs (with redaction boxes)
4. Users can specify keywords to exclude from redaction using a **whitelist**.

## Installation

This tool is intended for **advanced users** who are comfortable with Python and setting up dependencies.

### Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

### Required Libraries

Install the required libraries listed in the `requirements.txt` file:
```bash
pip install -r [requirements.txt](http://_vscodecontentref_/2)
```

Additionally, download the spaCy language model:
```bash
python -m spacy download en_core_web_sm
```

## Usage

### Running the Script

1. Clone this repository or download the script.
2. Open a terminal and navigate to the script's directory.
3. Run the script:
   ```bash
   python doc_anonymizer.py
   ```

### Modes of Operation

1. **Single File Mode**:
   - Anonymize a single file by providing its path.
   - Specify the output folder for the anonymized file.

2. **Batch Folder Mode**:
   - Anonymize all supported files in a folder.
   - Specify the input folder and the output folder.

### Whitelist Customization

The **whitelist** feature allows users to specify keywords that should not be redacted. This can be customized based on your business or sector to ensure that important terms remain visible in the anonymized documents. For example:
- A marketing agency might whitelist client names or product names.
- A legal firm might whitelist specific legal terms or case references.

When prompted, simply enter a comma-separated (.csv file) list of keywords (e.g., `CompanyName, New York`) to exclude them from redaction.

### Output

- Anonymized files are saved in the specified output folder with `_anonymized` appended to the filename.
- A CSV log (`anonymization_log.csv`) is generated in the output folder, listing all redacted entities.

## Example

### Input Document (Word)
```
John Doe works at Microsoft in New York.
```

### Output Document
```
[PERSON] works at [ORG_NAME] in [LOCATION].
```

### Anonymization Log
| Timestamp           | Filename                  | Original       | Replacement   |
|---------------------|---------------------------|----------------|---------------|
| 2025-04-14 12:00:00 | example.docx              | John Doe       | [PERSON]      |
| 2025-04-14 12:00:00 | example.docx              | Microsoft      | [ORG_NAME]    |
| 2025-04-14 12:00:00 | example.docx              | New York       | [LOCATION]    |

## Limitations

- **Footnotes**: The tool currently has limited support for anonymizing footnotes in Word documents.
- **PDF Redaction**: Redaction in PDFs is limited to searchable text. Images or scanned PDFs are not supported.

### Future Plans

- **Improved PII Detection**:
  - Enhance the PII detection model to increase accuracy and reduce false positives.
  - Explore integrating advanced NLP models or custom-trained models tailored to specific industries.
- **Narrowed Redactions**:
  - Implement redaction rules based on related client names, predefined country lists, or industry-specific terms to provide more context-aware anonymization.
- **Enhanced Footnote Support**:
  - Improve the handling of footnotes in Word documents to ensure complete anonymization.
- **Scanned PDF Support**:
  - Add OCR (Optical Character Recognition) capabilities to process and redact text in scanned PDFs.

## Contributing

Contributions are welcome! If you have suggestions for improvements or new features, feel free to open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).

## Acknowledgments

- Built with [spaCy](https://spacy.io/) for natural language processing.
- Inspired by the challenges faced by freelancers and solopreneurs in protecting sensitive information.