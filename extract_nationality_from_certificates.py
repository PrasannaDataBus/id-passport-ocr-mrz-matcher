"""
======================
NATIONALITY EXTRACTION PIPELINE (OCR + MRZ)
======================

Purpose:
--------
To extract the nationality of a person from scanned certificates, passports, or visas.
This script combines OCR (Tesseract), MRZ reading (passporteye), and robust
text analysis to identify the most likely country of nationality — even
when the text is partially distorted or multilingual (English + Arabic).

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

import os
import pandas as pd
import re

from termcolor import colored
import openpyxl
from tabulate import tabulate as tab

# Load the CSV
csv_path = r"C:\xxx\yyy\aaa\Certificate.csv"
df = pd.read_csv(csv_path)


# Function to extract only the first number sequence after "-"
def extract_number(value):
    if pd.isna(value):
        return None
    value = os.path.splitext(str(value))[0]  # remove extension
    nums = re.findall(r'\d+', value)
    if len(nums) >= 2:
        return nums[-1]   # take the last group
    elif len(nums) == 1:
        return nums[0]    # only one number group
    return None


# Apply cleaning logic to CSV column
df['cleaned_img_certificate'] = df['img_certificate'].apply(extract_number)
print(colored("CSV cleaned preview:", 'red'))
df.info()
print(tab(df.head(5), headers='keys', tablefmt='psql', showindex=False))

# Process Certificate Folder

folder_path = r"C:\xxx\yyy\aaa\certificate"

file_data = []
for file in os.listdir(folder_path):
    if file.lower().endswith(('.jpg', '.jpeg', '.png', '.pdf')):
        full_path = os.path.join(folder_path, file)
        raw_name = os.path.splitext(file)[0]   # filename without extension
        cleaned = extract_number(raw_name)
        file_data.append([raw_name, file, cleaned])

df_files = pd.DataFrame(file_data, columns=['ftp_img_certificate', 'ftp_real_img_certificate', 'ftp_cleaned_img_certificate'])
print(colored("\nFolder file table preview:", 'red'))
df_files.info()
print(tab(df_files.head(5), headers='keys', tablefmt='psql', showindex=False))

# Compare and Keep Matches
merged_df = pd.merge(df_files, df, left_on='ftp_cleaned_img_certificate', right_on='cleaned_img_certificate', how='inner')

# Keep only one match per certificate ID
merged_unique = merged_df.drop_duplicates(subset="cleaned_img_certificate", keep="first")

print(colored("\nMerged result preview:", 'red'))
merged_unique.info()
print(tab(merged_unique.head(25), headers='keys', tablefmt='psql', showindex=False))

import pycountry
from passporteye import read_mrz
import pytesseract

from pdf2image import convert_from_path
from PIL import Image, UnidentifiedImageError, ImageFile
from difflib import get_close_matches

import cv2

######################################### SYSTEM PATH CONFIGURATION ####################################################

# Tesseract OCR executable path (change if installed elsewhere)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Poppler library path for PDF conversion
POPPLER_PATH = r"C:\xxxx\yyyy\dddd\poppler-25.07.0\Library\bin"

########################################### COUNTRY & LANGUAGE MAPPINGS ################################################

# ISO & dictionaries
ISO3_TO_NAME = {c.alpha_3: c.name for c in pycountry.countries}
ISO3_SET = set(ISO3_TO_NAME.keys())
BLOCKLIST_ISO3 = {"ATA"}  # Exclude invalid or irrelevant codes (like Antarctica)

# Common OCR confusions or shorthand → corrected ISO3 codes
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

# Arabic text → ISO3 codes (for bilingual documents)
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

############################################# IMAGE HANDLING & OCR PREP ################################################

# Prevent Pillow warnings on large files or truncated images

import imghdr

# Safe Image Loader
ImageFile.LOAD_TRUNCATED_IMAGES = True
Image.MAX_IMAGE_PIXELS = None  # disable decompression bomb warning safely

try:
    import pillow_heif
    pillow_heif.register_heif_opener()
except ImportError:
    pass  # optional


# helpers


def to_country_name(iso3):
    """Return full country name from ISO3 code (e.g. 'IND' → 'India')."""
    if not iso3 or iso3 in BLOCKLIST_ISO3: return None
    return ISO3_TO_NAME.get(iso3)


def img_from_file(file_path):
    """Return PIL image from image or first page of PDF, with real format detection.
    Safely open an image or first page of a PDF.Converts all images to RGB to ensure consistent OCR input"""
    try:
        # Handle PDF
        if file_path.lower().endswith('.pdf'):
            pages = convert_from_path(file_path, dpi=250, poppler_path=POPPLER_PATH)
            return pages[0] if pages else None

        # Verify real image type
        real_type = imghdr.what(file_path)
        if real_type is None:
            print("Unsupported or mislabeled file:", file_path)
            return None

        # Open normally
        with Image.open(file_path) as im:
            im = im.convert("RGB")
        return im

    except UnidentifiedImageError:
        print("Unidentified/corrupted image:", file_path)
        return None
    except Exception as e:
        print("Image load failed:", file_path, "→", e)
        return None


def preprocess_for_ocr(pil_img):
    """PIL → OpenCV preprocess → PIL back (improves OCR).
        Preprocess an image for better OCR accuracy.
        - Convert to grayscale
        - Enhance contrast (CLAHE)
        - Apply Otsu’s thresholding
        """
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
    """Run OCR using Tesseract in English + Arabic."""
    return pytesseract.image_to_string(pil_img, lang="eng+ara")


def find_after_keyword(text_up, keywords):
    """Return token(s) that follow a keyword on the same line.
    Find words appearing immediately after known keywords like 'NATIONALITY' or 'الجنسية'.
    Useful when nationality appears on same line as the label."""
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
    """Match full official country names (EN/AR) in OCR text."""
    # English names
    for iso, name in ISO3_TO_NAME.items():
        if name.upper() in text_up and iso not in BLOCKLIST_ISO3:
            return iso
    # Arabic common names
    for ar, iso in ARABIC_NAME_TO_ISO3.items():
        if ar in text_up:
            return iso
    return None


def extract_country_word(text_up):
    """
    Robustly detect full country names or adjectives in OCR text (e.g. 'INDIA', 'EGYPTIAN', 'UNITED ARAB EMIRATES').
    Works in English and Arabic.
    Detect full country words, adjectives, or Arabic equivalents in text.
    Corrects OCR errors (e.g. INDlA → INDIA) and handles demonyms (“EGYPTIAN”, “FILIPINO”).
    """
    text_up = re.sub(r"[^A-Z\s]", " ", text_up.upper())

    # Common variants / adjectives / frequent OCR typos
    variants = {
        "INDIA": "IND", "INDlA": "IND", "INDIAN": "IND",
        "UNITED ARAB EMIRATES": "ARE", "EMIRATES": "ARE", "EMIRATI": "ARE", "UAE": "ARE",
        "PAKISTAN": "PAK", "PAKISTANI": "PAK", "PAKlSTAN": "PAK",
        "EGYPT": "EGY", "EGYPTIAN": "EGY",
        "LEBANON": "LBN", "LEBANESE": "LBN",
        "SYRIA": "SYR", "SYRIAN": "SYR",
        "JORDAN": "JOR", "JORDANIAN": "JOR",
        "IRAQ": "IRQ", "IRAQI": "IRQ",
        "IRAN": "IRN", "IRANIAN": "IRN",
        "OMAN": "OMN", "OMANI": "OMN",
        "QATAR": "QAT", "QATARI": "QAT",
        "BAHRAIN": "BHR", "BAHRAINI": "BHR",
        "KUWAIT": "KWT", "KUWAITI": "KWT", "KWI": "KWT",
        "PHILIPPINES": "PHL", "FILIPINO": "PHL",
        "BANGLADESH": "BGD", "BANGLADESHI": "BGD",
        "NEPAL": "NPL", "NEPALI": "NPL",
        "SRI LANKA": "LKA", "SRILANKA": "LKA", "SRILANKAN": "LKA",
        "CHINA": "CHN", "CHINESE": "CHN",
        "NIGERIA": "NGA", "NIGERIAN": "NGA",
        "KENYA": "KEN", "KENYAN": "KEN",
        "ETHIOPIA": "ETH", "ETHIOPIAN": "ETH",
        "SOUTH AFRICA": "ZAF", "SOUTH AFRICAN": "ZAF",
        "TURKEY": "TUR", "TURKISH": "TUR",
        "SAUDI ARABIA": "SAU", "SAUDI": "SAU", "KSA": "SAU",
        "PALESTINE": "PSE", "PALESTINIAN": "PSE",
        "LIBYA": "LBY", "LIBYAN": "LBY",
        "MOROCCO": "MAR", "MOROCCAN": "MAR",
        "TUNISIA": "TUN", "TUNISIAN": "TUN",
        "ALGERIA": "DZA", "ALGERIAN": "DZA",
        "FRANCE": "FRA", "FRENCH": "FRA",
        "GERMANY": "DEU", "GERMAN": "DEU",
        "UNITED KINGDOM": "GBR", "ENGLAND": "GBR", "BRITISH": "GBR",
        "USA": "USA", "UNITED STATES": "USA", "AMERICAN": "USA",
        "CANADA": "CAN", "CANADIAN": "CAN",
        "ITALY": "ITA", "ITALIAN": "ITA",
        "SPAIN": "ESP", "SPANISH": "ESP",
        "PORTUGAL": "PRT", "PORTUGUESE": "PRT",
        "RUSSIA": "RUS", "RUSSIAN": "RUS",
        "UKRAINE": "UKR", "UKRAINIAN": "UKR",
        "POLAND": "POL", "POLISH": "POL",
        "SWITZERLAND": "CHE", "SWISS": "CHE",
        "NETHERLANDS": "NLD", "DUTCH": "NLD", "HOLLAND": "NLD",
        "BELGIUM": "BEL", "BELGIAN": "BEL",
        "SINGAPORE": "SGP", "SINGAPOREAN": "SGP",
        "MALAYSIA": "MYS", "MALAYSIAN": "MYS",
        "THAILAND": "THA", "THAI": "THA",
        "INDONESIA": "IDN", "INDONESIAN": "IDN",
        "JAPAN": "JPN", "JAPANESE": "JPN",
        "KOREA": "KOR", "KOREAN": "KOR",
        "AFGHANISTAN": "AFG", "AFGHAN": "AFG"
    }

    # Arabic variants (common in Gulf IDs)
    arabic_variants = {
        "الهند": "IND", "الإمارات": "ARE", "الامارات": "ARE", "السعودية": "SAU",
        "قطر": "QAT", "البحرين": "BHR", "عمان": "OMN", "مصر": "EGY",
        "الأردن": "JOR", "الاردن": "JOR", "سوريا": "SYR", "لبنان": "LBN",
        "العراق": "IRQ", "إيران": "IRN", "ايران": "IRN", "فلسطين": "PSE",
        "الفلبين": "PHL", "باكستان": "PAK", "نيبال": "NPL", "بنغلاديش": "BGD",
        "سريلانكا": "LKA", "الصين": "CHN", "نيجيريا": "NGA", "اثيوبيا": "ETH",
        "جنوب افريقيا": "ZAF"
    }

    # Combine all
    combined = {**variants, **arabic_variants}

    for key, iso in combined.items():
        if key in text_up:
            return iso

    # fallback: full official name
    for iso, name in ISO3_TO_NAME.items():
        if name.upper() in text_up:
            return iso

    return None


# optional, only if needed later
def fuzzy_iso3_safe(code):
    """Fuzzy match slightly incorrect 3-letter codes (e.g. INO → IND)."""
    if not code or not re.fullmatch(r"[A-Z]{3}", code): return None
    hit = get_close_matches(code, list(ISO3_SET), n=1, cutoff=0.94)
    if hit and hit[0] not in BLOCKLIST_ISO3:
        return hit[0]
    return None


############################################ MAIN NATIONALITY EXTRACTOR ###############################################

# main extractor returning rich info


def extract_nationality_best(file_path):
    """
        Unified nationality extraction from image or PDF.
        Combines MRZ reading and OCR-based multi-step analysis.
        """
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

        # 2d) Full word scan (e.g. INDIA, PAKISTAN, EGYPT, etc.)
        iso = extract_country_word(text_up)
        if iso:
            return {
                "nationality_raw": ISO3_TO_NAME[iso],
                "nationality_iso3": iso,
                "nationality_country": to_country_name(iso),
                "nationality_confidence": 0.9,
                "nationality_source": "FULL_WORD_SCAN"
            }

        # 2e) ultra-conservative fuzzy on 3-letter token
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

# Save both full and filtered outputs
merged_unique.to_csv(r"C:\xxxx\yyyy\dddd\matched_certificates_with_nationality_raw.csv", index=False)
good.to_csv(r"C:\xxxx\yyyy\dddd\matched_certificates_with_nationality_confident.csv", index=False)

print("\n Done! Matched file saved")
