# id-passport-ocr-mrz-matcher

A Python tool that matches certificate image/PDF files with records from a CSV file and automatically extracts **nationality information** using **OCR**, **MRZ**, and **ISO3 country code mapping**.

"""
======================
NATIONALITY EXTRACTION PIPELINE (OCR + MRZ)
======================

Purpose:
--------
To extract the nationality of a person from scanned certificates, passports, or visas.
This script combines OCR (Tesseract), MRZ reading (passporteye), and robust
Text analysis to identify the most likely country of nationality — even
When the text is partially distorted or multilingual (English + Arabic).

Core Steps:
------------
1. Try MRZ (Machine Readable Zone) extraction if available.
2. OCR the document image or PDF.
3. Detect nationality based on:
   - Keyword lines (“NATIONALITY: IND”)
   - 3-letter ISO3 codes
   - Full country names or adjectives (“INDIA”, “EGYPTIAN”)
   - Arabic text (e.g., “الهند”, “الإمارات”)
   - Fuzzy matching for misspelled 3-letter codes

Output:
--------
The pipeline returns:
- nationality_raw         → The raw extracted string (e.g., “IND”, “INDIA”, “الهند”)
- nationality_iso3        → Clean ISO-3 country code (e.g., “IND”)
- nationality_country     → Full country name (e.g., “India”)
- nationality_confidence  → Confidence score (0–1)
- nationality_source      → Which logic path found it (“MRZ”, “FULL_WORD_SCAN”, etc.)

Dependencies:
-------------
pip install passporteye pycountry pytesseract pdf2image pillow opencv-python
and install:
- Tesseract OCR
- Poppler (for PDF rendering)
"""

```bash
pip install pandas tabulate termcolor passporteye pytesseract pycountry pdf2image pillow opencv-python

Install the below in local

tesseract.exe
poppler-25.07.0
