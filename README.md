# id-passport-ocr-mrz-matcher

A Python tool that matches certificate image/PDF files with records from a CSV file and automatically extracts **nationality information** using **OCR**, **MRZ**, and **ISO3 country code mapping**.

## Features
- Matches certificates with CSV entries using numeric codes in filenames  
- Extracts nationality from certificates using:
  - MRZ parsing (passport-style documents)
  - OCR text recognition (English + Arabic)
  - Keyword-based detection
  - ISO3, alias, and Arabic country name mapping
- Generates confidence scores and identifies data source  
- Outputs both full and high-confidence result files  

## Setup
### 1. Install Dependencies

```bash
pip install pandas tabulate termcolor passporteye pytesseract pycountry pdf2image pillow opencv-python

Install the below in local

tesseract.exe
poppler-25.07.0
