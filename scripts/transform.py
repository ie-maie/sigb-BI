import pandas as pd
import re
import numpy as np

ARABIC_PATTERN = re.compile(r'[\u0600-\u06FF]')

EXPECTED_COLUMNS = [
    "Cote", "Titre", "Auteur", "Lieu", "Edition",
    "Annee", "Nb_pages", "Matiere", "Inventaire"
]

EXACT_ROLES = [
    r's/s\s*\.?\s*dir\.?', r'S/S\s*dir', r's/sdir', r's\.sdir',
    r'coord', r'cod', r'cor\.', r'Edited\s+by', r'published',
    r'\béd\b\.?', r'\bed\b\.?', r'\bèd\b\.?', r'\bdir\b', r'\bpub\b', r'\bpréf\b', r'\bpré\b', r'\bpref\b',
    r'إعداد', r'تنسيق', r'إشراف', r'تحت\s+اشراف', r'تحقيق', r'تحرير', r'تحقيق\s+ودراسة'
]

ROLE_REGEX = re.compile(r'[\(\[\{]\s*(?:' + '|'.join(EXACT_ROLES) + r')\s*[\)\]\}]', re.IGNORECASE)

CORPORATE_KEYWORDS = [
    'faculte', 'fac.', 'ministere', 'centre', 'association', 'royaume', 'comite', 'urbama', 'banque',
    'جامعة', 'مركز', 'وزارة', 'جمعية', 'لجنة', 'مؤسسة', 'منظمة', 'إدارة', 'اتحاد', 'إتحاد', 'البنك'
]

AUTHOR_SPLIT_REGEX = re.compile(
    r'(?:\s*;\s*|\s*--\s*|\s+et\s+|\s+and\s+|\s+و\s+|(?<=[^\b[A-Z]\b])\.\s+(?=[A-Z\u0600-\u06FF]))'
)


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = EXPECTED_COLUMNS
    return df


def clean_cote(df: pd.DataFrame) -> pd.DataFrame:
    if "Cote" in df.columns:
        # Fixed: Strips triple quotes and escaping anomalies coming from spreadsheet source exports
        df["Cote"] = df["Cote"].astype(str).str.replace('"', '', regex=False).str.strip()
        df["Cote"] = df["Cote"].replace(["nan", "", "None", "<NA>"], "Sans Cote")
    return df


def extract_roles_and_split_authors(author_str):
    if pd.isna(author_str) or str(author_str).lower() in ["nan", "", "none"]:
        return [("Auteur Inconnu", "Auteur")]

    author_str = str(author_str).strip()
    detected_role = "Auteur"

    found_roles = ROLE_REGEX.findall(author_str)
    if found_roles:
        detected_role = re.sub(r'[\(\[\{\)\]\}]', '', found_roles[0]).strip()
        author_str = ROLE_REGEX.sub("", author_str)

    is_corporate = any(kw in author_str.lower() for kw in CORPORATE_KEYWORDS)

    if is_corporate:
        auth_cleaned = re.sub(r'^[.,\s\-\/]+|[.,\s\-\/]+$', '', author_str).strip()
        if auth_cleaned:
            return [(auth_cleaned, detected_role)]
        return [("Auteur Inconnu", "Auteur")]

    raw_authors = AUTHOR_SPLIT_REGEX.split(author_str)
    author_role_pairs = []
    for auth in raw_authors:
        auth_cleaned = re.sub(r'^[.,\s\(\)\[\]\-]+|[.,\s\(\)\[\]\-]+$', '', auth).strip()
        if auth_cleaned and auth_cleaned.lower() not in ["nan", "unknown", "none"]:
            author_role_pairs.append((auth_cleaned, detected_role))

    return author_role_pairs if author_role_pairs else [("Auteur Inconnu", "Auteur")]


def clean_text_and_extract(df: pd.DataFrame) -> pd.DataFrame:
    punctuation_sensitive_cols = ["Lieu", "Edition"]

    pairs_series = df["Auteur"].apply(extract_roles_and_split_authors)
    df["Auteur_Pairs"] = pairs_series
    df["Auteur_Names"] = pairs_series.apply(lambda pairs: "; ".join([p[0] for p in pairs]))

    for col in ["Titre", "Lieu", "Edition", "Matiere"]:
        df[col] = df[col].astype(str).str.strip()
        if col in punctuation_sensitive_cols:
            df[col] = df[col].str.replace(r'[.,\s]+$', '', regex=True)
        
        if col == "Edition":
            df[col] = df[col].replace(["nan", "", "None", "<NA>"], "Édition Inconnue")
        elif col == "Titre":
            df[col] = df[col].replace(["nan", "", "None", "<NA>"], "Titre Inconnu")
        else:
            df[col] = df[col].replace(["nan", "", "None", "<NA>"], None)

    return df


def convert_types(df: pd.DataFrame) -> pd.DataFrame:
    for col in ["Annee", "Nb_pages"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
        df[col] = df[col].astype("Int64")
    return df


def add_language(df: pd.DataFrame, lang_code: str) -> pd.DataFrame:
    df["language"] = lang_code
    return df


def create_notice_key(df: pd.DataFrame) -> pd.DataFrame:
    df["notice_key"] = (
        df["Titre"].fillna("Titre Inconnu") + "|" +
        df["Auteur_Names"].fillna("Auteur Inconnu") + "|" +
        df["Annee"].astype(str).replace('<NA>', '')
    )
    # Uniform tracking stabilization
    df["notice_key"] = df["notice_key"].astype(str).str.strip()
    return df


def handle_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    temp_df = df.drop(columns=["Auteur_Pairs"])
    df["is_duplicate_row"] = temp_df.duplicated(keep="first")
    df["is_duplicate_notice"] = temp_df.duplicated("notice_key", keep="first")
    return df


def transform(df: pd.DataFrame, lang_code: str) -> pd.DataFrame:
    df = standardize_columns(df)
    df = clean_cote(df)
    df = clean_text_and_extract(df)
    df = convert_types(df)
    df = add_language(df, lang_code)
    df = create_notice_key(df)
    df = handle_duplicates(df)

    df["Auteur"] = df["Auteur_Names"]

    # Only drop exact layout duplicates. Multiple items for a notice remain intact here.
    df = df[df["is_duplicate_row"] == False]
    return df.reset_index(drop=True)