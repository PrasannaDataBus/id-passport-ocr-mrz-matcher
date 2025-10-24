# id-passport-ocr-mrz-matcher

# Certificate Matcher

This project matches certificate image/PDF files with entries from a CSV file by extracting the numeric code embedded in filenames. It also attempts to automatically detect **nationality** from the certificate using **MRZ (passport)** data or **OCR text scanning**.

## Features
- Extracts certificate ID numbers from filenames
- Matches certificate images to CSV records
- Uses MRZ (machine-readable passport zone) when available
- Falls back to OCR text recognition (English + Arabic)
- Attempts to detect nationality using standard ISO country codes
- Outputs a clean merged dataset with nationality information

## Requirements
Install dependencies:

```bash
pip install pandas tabulate termcolor passporteye pytesseract pycountry pdf2image pillow
