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
from difflib import get_close_matches
import cv2

# Windows: Path to tesseract.exe
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Windows: Poppler bin folder path
POPPLER_PATH = r"C:\zzz\xxxx\yyyy\poppler-25.07.0\Library\bin"

# ISO & dictionaries
ISO3_TO_NAME = {c.alpha_3: c.name for c in pycountry.countries}
ISO3_SET = set(ISO3_TO_NAME.keys())
BLOCKLIST_ISO3 = {"ATA"}  # Antarctica

# Common aliases / OCR confusions → ISO3
ALIAS_TO_ISO3 = {
    "LEB":"LBN",
    "IRA":"IRN",
    "IRN":"IRN",
    "IRQ":"IRQ",
    "SYR":"SYR",
    "EGY":"EGY",
    "JOR":"JOR",
    "UAE":"ARE",
    "ARE":"ARE",
    "KSA":"SAU",
    "SAU":"SAU",
    "KWT":"KWT",
    "KWI":"KWT",
    "QAT":"QAT",
    "BHR":"BHR",
    "OMN":"OMN",
    "PHI":"PHL",
    "PHL":"PHL",
    "PAK":"PAK",
    "IND":"IND",
    "BGD":"BGD",
    "NPL":"NPL",
    "LKA":"LKA",
    "NIG":"NGA",
    "NGA":"NGA",
    "ETH":"ETH",
    "KEN":"KEN",
    "ZAF":"ZAF",
    "CHN":"CHN",
    "HKG":"HKG",
    "KOR":"KOR",
    "PRK":"PRK",
    "POR":"PRT",
    "PRT":"PRT",
    "ROM":"ROU",
    "ROU":"ROU",
    "GER":"DEU",
    "DEU":"DEU",
    "SWI":"CHE",
    "CHE":"CHE",
    "NED":"NLD",
    "NET":"NLD",
    "HOL":"NLD",
    "NLD":"NLD",
    "RUS":"RUS",
    "USA":"USA",
    "GBR":"GBR",
    "UK":"GBR",
    "TUR":"TUR",
    "ALG":"DZA",
    "DZA":"DZA",
    "SUD":"SDN",
    "SDN":"SDN",
    "TUN":"TUN",
    "MAR":"MAR",
    "LIB":"LBY",
    "LBY":"LBY",
    "PAL":"PSE",
    "PSE":"PSE",
    "KUW":"KWT",
    "EMI":"ARE"
}

# Arabic country name (common OCR) → ISO3 (expand as you encounter more)
ARABIC_NAME_TO_ISO3 = {
    "الكويت":"KWT",
    "الإمارات":"ARE",
    "الامارات":"ARE",
    "السعودية":"SAU",
    "قطر":"QAT",
    "البحرين":"BHR",
    "عمان":"OMN",
    "مصر":"EGY",
    "الأردن":"JOR",
    "الاردن":"JOR",
    "سوريا":"SYR",
    "لبنان":"LBN",
    "العراق":"IRQ",
    "إيران":"IRN",
    "ايران":"IRN",
    "فلسطين":"PSE",
    "الهند":"IND",
    "باكستان":"PAK",
    "الفلبين":"PHL",
    "بنغلاديش":"BGD",
    "نيبال":"NPL",
    "سريلانكا":"LKA",
    "الصين":"CHN",
    "نيجيريا":"NGA",
    "اثيوبيا":"ETH",
    "إثيوبيا":"ETH",
    "جنوب افريقيا":"ZAF"
}


# helpers
def to_country_name(iso3):
    if not iso3 or iso3 in BLOCKLIST_ISO3: return None
    return ISO3_TO_NAME.get(iso3)


def img_from_file(file_path):
    """Load first page image for PDFs; PIL Image for images."""
    if file_path.lower().endswith(".pdf"):
        pages = convert_from_path(file_path, dpi=350, poppler_path=POPPLER_PATH)
        if not pages: return None
        return pages[0]
    else:
        return Image.open(file_path)


def preprocess_for_ocr(pil_img):
    """PIL → OpenCV preprocess → PIL back (improves OCR)."""
    try:
        cv = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    except Exception:
        cv = cv2.imread(pil_img) if isinstance(pil_img, str) else None
    if cv is None: return pil_img

    gray = cv2.cvtColor(cv, cv2.COLOR_BGR2GRAY)
    # CLAHE for contrast
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8)).apply(gray)
    # Otsu threshold
    th = cv2.threshold(clahe, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    # Convert back to PIL
    return Image.fromarray(th)


def extract_text(pil_img):
    return pytesseract.image_to_string(pil_img, lang="eng+ara")


def find_after_keyword(text_up, keywords):
    """Return token(s) that follow a keyword on the same line."""
    lines = [ln.strip() for ln in text_up.splitlines() if ln.strip()]
    for ln in lines:
        for kw in keywords:
            if kw in ln:
                # take everything after kw
                tail = ln.split(kw, 1)[-1].strip(" :\t-")
                # first 1–2 tokens
                tokens = re.findall(r"[A-Z]{2,}|[A-Z]{3}", tail)
                if tokens:
                    return tokens[:2]
    return None


def strict_iso3_in_text(text_up):
    """Return first valid ISO3 code with word boundaries."""
    for m in re.finditer(r"\b[A-Z]{3}\b", text_up):
        code = m.group(0)
        # allow direct ISO or alias mapped to ISO
        if code in ISO3_SET and code not in BLOCKLIST_ISO3:
            return code
        if code in ALIAS_TO_ISO3:
            iso = ALIAS_TO_ISO3[code]
            if iso in ISO3_SET and iso not in BLOCKLIST_ISO3:
                return iso
    return None


def name_match_in_text(text_up):
    # English names
    for iso, name in ISO3_TO_NAME.items():
        if name.upper() in text_up and iso not in BLOCKLIST_ISO3:
            return iso
    # Arabic common names
    for ar, iso in ARABIC_NAME_TO_ISO3.items():
        if ar in text_up:
            return iso
    return None


# optional, only if needed later
def fuzzy_iso3_safe(code):
    if not code or not re.fullmatch(r"[A-Z]{3}", code): return None
    hit = get_close_matches(code, list(ISO3_SET), n=1, cutoff=0.94)
    if hit and hit[0] not in BLOCKLIST_ISO3:
        return hit[0]
    return None

# main extractor returning rich info


def extract_nationality_best(file_path):
    # 1) MRZ
    try:
        mrz = read_mrz(file_path)
        if mrz:
            nat = (mrz.to_dict() or {}).get("nationality")
            if nat and nat.upper() in ISO3_SET and nat.upper() not in BLOCKLIST_ISO3:
                iso = nat.upper()
                return {
                    "nationality_raw": iso,
                    "nationality_iso3": iso,
                    "nationality_country": to_country_name(iso),
                    "nationality_confidence": 1.0,
                    "nationality_source": "MRZ"
                }
    except Exception:
        pass

    # 2) OCR pipeline
    try:
        pil = img_from_file(file_path)
        if pil is None:
            return None
        pil_p = preprocess_for_ocr(pil)
        text = extract_text(pil_p)
        text_up = text.upper()

        # 2a) Keyword anchored (English & Arabic)
        tokens = find_after_keyword(text_up, keywords=["NATIONALITY", "الجنسية"])
        if tokens:
            # try first token as ISO3 or alias
            for t in tokens:
                t3 = re.sub(r"[^A-Z]", "", t)[:3]
                if len(t3) == 3:
                    if t3 in ISO3_SET and t3 not in BLOCKLIST_ISO3:
                        iso = t3
                        return {"nationality_raw": t,
                                "nationality_iso3": iso,
                                "nationality_country": to_country_name(iso),
                                "nationality_confidence": 0.9,
                                "nationality_source": "KEYWORD_LINE_ISO3"}
                    if t3 in ALIAS_TO_ISO3:
                        iso = ALIAS_TO_ISO3[t3]
                        return {"nationality_raw": t,
                                "nationality_iso3": iso,
                                "nationality_country": to_country_name(iso),
                                "nationality_confidence": 0.9,
                                "nationality_source": "KEYWORD_LINE_ALIAS"}
            # if not ISO3, try name near keyword
            iso = name_match_in_text(" ".join(tokens))
            if iso:
                return {"nationality_raw": " ".join(tokens),
                        "nationality_iso3": iso,
                        "nationality_country": to_country_name(iso),
                        "nationality_confidence": 0.85,
                        "nationality_source": "KEYWORD_LINE_NAME"}

        # 2b) Strict ISO3 scan anywhere
        iso = strict_iso3_in_text(text_up)
        if iso:
            return {"nationality_raw": iso,
                    "nationality_iso3": iso,
                    "nationality_country": to_country_name(iso),
                    "nationality_confidence": 0.85,
                    "nationality_source": "ISO3_SCAN"}

        # 2c) Full name scan (EN/AR)
        iso = name_match_in_text(text_up)
        if iso:
            return {"nationality_raw": ISO3_TO_NAME[iso],
                    "nationality_iso3": iso,
                    "nationality_country": to_country_name(iso),
                    "nationality_confidence": 0.8,
                    "nationality_source": "NAME_SCAN"}

        # 2d) (Optional) ultra-conservative fuzzy on 3-letter token
        m = re.search(r"\b[A-Z]{3}\b", text_up)
        if m:
            guess = fuzzy_iso3_safe(m.group(0))
            if guess:
                return {"nationality_raw": m.group(0),
                        "nationality_iso3": guess,
                        "nationality_country": to_country_name(guess),
                        "nationality_confidence": 0.75,
                        "nationality_source": "FUZZY_ISO3"}

        return None

    except Exception as e:
        print("OCR pipeline failed:", file_path, "→", e)
        return None


def run_extract(row):
    f = row['ftp_real_img_certificate']
    fp = os.path.join(folder_path, f)
    out = extract_nationality_best(fp)
    if not out:
        return pd.Series({
            "nationality_raw": None,
            "nationality_iso3": None,
            "nationality_country": None,
            "nationality_confidence": 0.0,
            "nationality_source": None,
        })
    return pd.Series(out)


merged_unique[[
    "nationality_raw",
    "nationality_iso3",
    "nationality_country",
    "nationality_confidence",
    "nationality_source"
]] = merged_unique.apply(run_extract, axis=1)

# Proper case country name, keep NaNs
merged_unique['nationality_country'] = merged_unique['nationality_country'].where(
    merged_unique['nationality_country'].notna(), None
)
merged_unique['nationality_country'] = merged_unique['nationality_country'].astype(str).str.title()

# keep only good ones (>=0.85) OR MRZ source
good = merged_unique[
    (merged_unique['nationality_confidence'] >= 0.85) |
    (merged_unique['nationality_source'] == "MRZ")
]

# STEP 5: SAVE RESULTS TO CSV
# Save both full and filtered outputs

merged_unique.to_csv(r"C:\zzz\xxx\yyy\matched_certificates_with_nationality_raw.csv", index=False)
good.to_csv(r"C:\zzz\xxx\yyy\matched_certificates_with_nationality_confident.csv", index=False)

print("\nDone! Matched file saved to:")
