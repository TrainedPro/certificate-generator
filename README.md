# PPTX Text Replacement and PDF Conversion

This script replaces placeholder text in a PPTX file and converts the modified file to PDF. It can optionally delete the intermediary PPTX file after conversion.

## Prerequisites

- Python 3.x
- LibreOffice (for converting PPTX to PDF)

## Setup

1. **Clone the Repository**

   ```bash
   git clone https://github.com/TrainedPro/certificate_generator.git
   cd certificate_generator
   ```

2. **Install Dependencies**

   Create a virtual environment and install the required libraries using `requirements.txt`.

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # On Windows use: .venv\Scripts\activate
   pip install -r requirements.txt
   ```

3. **Install LibreOffice**

   Make sure LibreOffice is installed on your system. It is required for converting PPTX files to PDF. 

   - **Ubuntu/Debian**:
     
     ```bash
     sudo apt-get install libreoffice
     ```

   - **Windows/Mac**:
     Download and install LibreOffice from [LibreOffice.org](https://www.libreoffice.org/download/download/).

## Usage

To replace placeholder text in a PPTX file and convert it to PDF, use the following command:

```bash
python main.py --pptx <path_to_pptx> --placeholder <placeholder_text> --replacement <replacement_text> [--delete-intermediary] [--output-dir <output_directory>]
```

### Command-Line Arguments

- `--pptx`: Path to the PPTX template file (default: `template.pptx`)
- `--placeholder`: Text placeholder to replace (default: `placeholder_text`)
- `--replacement`: Text to replace the placeholder with (default: `Replacement Text`)
- `--delete-intermediary`: Optional flag to delete the intermediary PPTX file after conversion (default: `False`)
- `--output-dir`: Directory for output files (default: `output`)

### Example

Replace the placeholder text "PLACEHOLDER" with "John Doe" in `template.pptx`, convert it to a PDF, and save the results in the `output` directory.

```bash
python main.py --pptx template.pptx --placeholder PLACEHOLDER --replacement "John Doe" --delete-intermediary --output-dir output
```

## Logging

The script logs detailed information and errors to `app.log` and the console.