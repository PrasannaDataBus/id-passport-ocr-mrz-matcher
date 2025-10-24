import os
import pandas as pd
import re

from termcolor import colored
import openpyxl
from tabulate import tabulate as tab

# STEP 1: LOAD & CLEAN CSV

# Path to the CSV file that contains certificate info
csv_path = r"C:\Users\xxx\yyy\ID_Proof.csv"
df = pd.read_csv(csv_path)


# Function to extract the numeric certificate ID from a filename-like string
# Example:
#    "CERT-IMG-00234.jpg" → "00234"
#    "abc-123-456" → "456"
def extract_number(value):
    if pd.isna(value):
        return None
    
    # Remove file extension if present (e.g., .jpg, .png)
    value = os.path.splitext(str(value))[0]

    # Find all groups of digits within the string
    nums = re.findall(r'\d+', value)

    # If multiple number sequences exist, the last one is assumed to be the ID
    if len(nums) >= 2:
        return nums[-1] # take the last group
    elif len(nums) == 1:
        return nums[0] # only one number group

    return None


# Create a new column in CSV dataframe with cleaned extracted certificate number
df['cleaned_img_certificate'] = df['img_certificate'].apply(extract_number)

print(colored("CSV cleaned preview:", 'red'))
df.info()
print(tab(df.head(5), headers='keys', tablefmt='psql', showindex=False))

# STEP 2: READ CERTIFICATE FOLDER

folder_path = r"C:\Users\xxx\yyy\cert"

file_data = []

# Loop through every file in the certificates folder
for file in os.listdir(folder_path):
    # Only consider image and PDF files
    if file.lower().endswith(('.jpg', '.jpeg', '.png', '.pdf')):
        full_path = os.path.join(folder_path, file)

        # Filename without extension
        raw_name = os.path.splitext(file)[0]

        # Extract the numeric certificate code from the filename
        cleaned = extract_number(raw_name)

        # Store original + cleaned values
        file_data.append([raw_name, file, cleaned])

# Convert folder scan results into dataframe
df_files = pd.DataFrame(file_data, columns=[
    'ftp_img_certificate',         # filename base
    'ftp_real_img_certificate',    # real filename including extension
    'ftp_cleaned_img_certificate'  # extracted numeric ID
])

print(colored("\nFolder file table preview:", 'red'))
df_files.info()
print(tab(df_files.head(5), headers='keys', tablefmt='psql', showindex=False))

# STEP 3: MATCH CSV ENTRIES WITH FOLDER FILES

# Merge on the cleaned numeric certificate code column
merged_df = pd.merge(
    df_files,
    df,
    left_on='ftp_cleaned_img_certificate',
    right_on='cleaned_img_certificate',
    how='inner'
)

# Avoid duplicate matches (if multiple images mapped to the same ID)
merged_unique = merged_df.drop_duplicates(subset="cleaned_img_certificate", keep="first")

print(colored("\nMerged result preview:", 'red'))
merged_unique.info()
print(tab(merged_unique.head(25), headers='keys', tablefmt='psql', showindex=False))

# STEP 4: EXTRACT NATIONALITY FROM CERT FILE

import pycountry
from passporteye import read_mrz
import pytesseract
from pdf2image import convert_from_path
from PIL import Image

# Set Tesseract OCR executable path (for Windows)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Poppler path (needed to convert PDF → Image in Windows)
POPPLER_PATH = r"C:\Users\xxxx\yyyy\poppler-25.07.0\Library\bin"


def extract_nationality(file_path):
    """
    Attempts to detect nationality from a certificate / passport image.
    Tries 3 stages:
      1) MRZ extraction (if passport type)
      2) OCR scanning of image for nationality labels
      3) Fallback: match any known country name from text
    """

    # METHOD 1: Try extracting MRZ (passport machine-code zone)
    try:
        mrz = read_mrz(file_path)
        if mrz:
            nat = mrz.to_dict().get("nationality")
            if nat:
                country = pycountry.countries.get(alpha_3=nat.upper())
                return country.name if country else nat.upper()
    except:
        pass

    # METHOD 2: Convert PDF → Image or load image
    try:
        if file_path.lower().endswith('.pdf'):
            pages = convert_from_path(file_path, dpi=300, poppler_path=POPPLER_PATH)
            if not pages:
                return None
            img = pages[0]
        else:
            img = Image.open(file_path).convert("RGB")
    except Exception as e:
        print("Image/PDF load failed:", file_path, "→", e)
        return None

    # METHOD 3: OCR text extraction
    try:
        text = pytesseract.image_to_string(img, lang="eng+ara").upper()
    except Exception as e:
        print("OCR failed:", file_path, "→", e)
        return None

    # Look for common nationality label formats
    patterns = [
        r"NATIONALITY[:\s]*([A-Z]{3})",
        r"الجنسية[:\s]*([A-Z]{3})",
        r"\b(CITIZEN|NATIONAL|NACIONALIDAD|NATIONALITÉ)[^\n]*\b([A-Z]{3})"
    ]

    for p in patterns:
        m = re.search(p, text)
        if m:
            # Extract last matching captured group
            code = m.group(len(m.groups()))
            country = pycountry.countries.get(alpha_3=code.upper())
            return country.name if country else code.upper()

    # METHOD 4: As fallback, search full country names in text
    for country in pycountry.countries:
        if country.name.upper() in text:
            return country.name

    return None


# Apply nationality extraction to matched certificate files
merged_unique['nationality'] = merged_unique['ftp_real_img_certificate'].apply(
    lambda f: extract_nationality(os.path.join(folder_path, f))
)

# Show nationality extraction sample output
print(merged_unique[['ftp_real_img_certificate', 'nationality']].head(30))

# STEP 5: SAVE RESULTS TO CSV

output_path = r"C:\Users\xxxx\yyyy\nationality_for_analysis.csv"
merged_unique.to_csv(output_path, index=False)

print("\nDone! Matched file saved to:", output_path)
