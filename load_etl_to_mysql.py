"""load_etl_to_mysql.py

Loads ONLY the two cleaned CSVs produced by etl.py:
- buf_clean.csv
- bua_clean.csv

It does a full refresh: TRUNCATE target tables then reload.

Foreign keys are resolved by looking up IDs by their natural keys
(code_langue, nom_editeur, nom_complet, cote, libelle, num_inventaire).

Note: If your DB schema name is different from the default (or if some FK
tables are missing), truncation skips missing tables.
"""

import os
from typing import Optional, Dict, Tuple

import mysql.connector
from mysql.connector import Error
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

DB_HOST = os.getenv("DB_HOST")
DB_NAME = os.getenv("DB_NAME")
# Fallbacks if your .env points to a schema where tables are not present.
# Adjust if needed.
SCHEMA_CANDIDATES = [DB_NAME, "sigb", "SIGB"]

DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")

BUA_CSV = "bua_clean.csv"
BUF_CSV = "buf_clean.csv"

TARGET_TABLES = [
    "notice_auteur",
    "notice_matiere",
    "exemplaire",
    "notice",
    "matiere",
    "classification",
    "auteur",
    "editeur",
    "langue",
]

EXPECTED_NOTICE_COLS = {
    "Cote",
    "Titre",
    "Auteur",
    "Lieu",
    "Edition",
    "Annee",
    "Nb_pages",
    "Matiere",
    "Inventaire",
    "notice_key",
}


def get_connection():
    try:
        return mysql.connector.connect(
            host=DB_HOST,
            database=DB_NAME,
            user=DB_USER,
            password=DB_PASSWORD,
            charset="utf8mb4",
            use_unicode=True,
        )
    except Error as e:
        print(f"Error connecting to MySQL: {e}")
        return None


def truncate_tables(conn, cursor):
    print("\n=== TRUNCATING target tables (full reload) ===")

    cursor.execute("SHOW TABLES;")
    existing = {row[0] for row in cursor.fetchall()}

    for table in TARGET_TABLES:
        if table in existing:
            cursor.execute(f"TRUNCATE TABLE {table};")
        else:
            print(f"⚠️  Skip TRUNCATE (missing table): {table}")

    conn.commit()
    print("✅ Truncate complete")


def validate_columns(df: pd.DataFrame, expected: set, label: str):
    missing = expected - set(df.columns)
    if missing:
        raise RuntimeError(f"{label}: CSV missing columns: {sorted(missing)}")


def upsert_lookup_ids(
    cursor,
    table: str,
    value_col: str,
    values: pd.Series,
    insert_sql: str,
    select_sql: str,
) -> Dict[str, int]:
    """Insert distinct values into lookup table, then return mapping value->id."""

    distinct = (
        values.dropna()
        .astype(str)
        .map(lambda s: s.strip())
        .loc[lambda s: s != ""]
        .drop_duplicates()
    )

    if len(distinct) == 0:
        return {}

    print(f"Loading lookup {table}: {len(distinct)} distinct")

    for v in distinct:
        cursor.execute(insert_sql, (v,))

    cursor.execute(select_sql)
    rows = cursor.fetchall()

    # select_sql should return (id, natural_key_value)
    return {str(row[1]): row[0] for row in rows}


def load_two_csvs() -> Tuple[pd.DataFrame, pd.DataFrame]:
    if not os.path.exists(BUF_CSV):
        raise FileNotFoundError(f"Missing file: {BUF_CSV}")
    if not os.path.exists(BUA_CSV):
        raise FileNotFoundError(f"Missing file: {BUA_CSV}")

    buf = pd.read_csv(BUF_CSV, encoding="utf-8-sig")
    bua = pd.read_csv(BUA_CSV, encoding="utf-8-sig")

    validate_columns(buf, EXPECTED_NOTICE_COLS, "buf_clean.csv")
    validate_columns(bua, EXPECTED_NOTICE_COLS, "bua_clean.csv")

    return buf, bua


def full_refresh_load(conn, cursor, buf: pd.DataFrame, bua: pd.DataFrame):
    truncate_tables(conn, cursor)

    # Ensure langue rows exist (natural keys: code_langue)
    cursor.execute(
        "INSERT INTO langue (code_langue, libelle) VALUES (%s,%s),(%s,%s) "
        "ON DUPLICATE KEY UPDATE code_langue=VALUES(code_langue)",
        ("fre", "Français", "ara", "Arabe"),
    )
    conn.commit()

    # Combine for lookup discovery
    df_all = pd.concat([buf, bua], ignore_index=True)

    # Lookup tables
    editeur_map = upsert_lookup_ids(
        cursor=cursor,
        table="editeur",
        value_col="nom_editeur",
        values=df_all["Edition"],
        insert_sql="INSERT IGNORE INTO editeur (nom_editeur) VALUES (%s)",
        select_sql="SELECT id_editeur, nom_editeur FROM editeur",
    )

    auteur_map = upsert_lookup_ids(
        cursor=cursor,
        table="auteur",
        value_col="nom_complet",
        values=df_all["Auteur"],
        insert_sql="INSERT IGNORE INTO auteur (nom_complet) VALUES (%s)",
        select_sql="SELECT id_auteur, nom_complet FROM auteur",
    )

    classification_map = upsert_lookup_ids(
        cursor=cursor,
        table="classification",
        value_col="cote",
        values=df_all["Cote"],
        insert_sql="INSERT IGNORE INTO classification (cote) VALUES (%s)",
        select_sql="SELECT id_classification, cote FROM classification",
    )

    matiere_map = upsert_lookup_ids(
        cursor=cursor,
        table="matiere",
        value_col="libelle",
        values=df_all["Matiere"],
        insert_sql="INSERT IGNORE INTO matiere (libelle) VALUES (%s)",
        select_sql="SELECT id_matiere, libelle FROM matiere",
    )

    cursor.execute("SELECT id_langue, code_langue FROM langue")
    langue_map = {row[1]: row[0] for row in cursor.fetchall()}

    print("\n=== Loading notice + bridges ===")
    notice_key_to_id: Dict[str, int] = {}

    def safe_int_nullable(x) -> Optional[int]:
        if pd.isna(x):
            return None
        s = str(x).strip()
        if s == "":
            return None
        try:
            return int(float(s))
        except ValueError:
            return None

    def insert_notice_rows(df: pd.DataFrame, code_langue: str):
        nonlocal notice_key_to_id

        id_langue = langue_map.get(code_langue)
        if id_langue is None:
            raise RuntimeError(f"Missing langue id for {code_langue}")

        inserted = 0
        for _, r in df.iterrows():
            titre = str(r.get("Titre") or "").strip()
            if not titre:
                continue

            cote = str(r.get("Cote") or "").strip()
            edition = str(r.get("Edition") or "").strip()
            auteur = str(r.get("Auteur") or "").strip()
            lieu = str(r.get("Lieu") or "").strip()
            matiere = str(r.get("Matiere") or "").strip()
            note = None

            id_editeur = editeur_map.get(edition)
            id_classif = classification_map.get(cote)

            annee = safe_int_nullable(r.get("Annee"))
            nb_pages = safe_int_nullable(r.get("Nb_pages"))

            if id_editeur is None or id_classif is None:
                continue

            cursor.execute(
                """
                INSERT INTO notice
                    (titre, annee_pub, nb_pages, lieu_edition, id_editeur, id_langue, id_classification, note)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
                """,
                (
                    titre,
                    annee,
                    nb_pages,
                    lieu if lieu != "" else None,
                    id_editeur,
                    id_langue,
                    id_classif,
                    note,
                ),
            )
            id_notice = cursor.lastrowid
            notice_key_to_id[str(r["notice_key"])] = id_notice
            inserted += 1

        conn.commit()
        print(f"✅ inserted notices for {code_langue}: {inserted}")

    def insert_bridge_rows(df: pd.DataFrame):
        inserted_auteur = 0
        inserted_matiere = 0
        inserted_exemplaire = 0

        for _, r in df.iterrows():
            nk = str(r.get("notice_key") or "").strip()
            if not nk:
                continue

            id_notice = notice_key_to_id.get(nk)
            if not id_notice:
                continue

            auteur = str(r.get("Auteur") or "").strip()
            id_auteur = auteur_map.get(auteur) if auteur else None
            if id_auteur:
                cursor.execute(
                    "INSERT IGNORE INTO notice_auteur (id_notice, id_auteur) VALUES (%s,%s)",
                    (id_notice, id_auteur),
                )
                inserted_auteur += 1

            matiere = str(r.get("Matiere") or "").strip()
            id_mat = matiere_map.get(matiere) if matiere else None
            if id_mat:
                cursor.execute(
                    "INSERT IGNORE INTO notice_matiere (id_notice, id_matiere) VALUES (%s,%s)",
                    (id_notice, id_mat),
                )
                inserted_matiere += 1

            inv = r.get("Inventaire")
            if pd.notna(inv):
                inv_str = str(inv).strip()
                if inv_str:
                    cursor.execute(
                        """
                        INSERT IGNORE INTO exemplaire (id_notice, num_inventaire, etat)
                        VALUES (%s,%s,%s)
                        """,
                        (id_notice, inv_str, "Disponible"),
                    )
                    inserted_exemplaire += 1

        conn.commit()
        print(
            "✅ bridges inserted: "
            f"notice_auteur={inserted_auteur}, notice_matiere={inserted_matiere}, exemplaire={inserted_exemplaire}"
        )

    insert_notice_rows(buf, "fre")
    insert_notice_rows(bua, "ara")

    insert_bridge_rows(buf)
    insert_bridge_rows(bua)


def print_summary(cursor):
    cursor.execute("SELECT COUNT(*) FROM notice")
    notice_count = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM exemplaire")
    exemplaire_count = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM auteur")
    auteur_count = cursor.fetchone()[0]

    print("\nData Summary:")
    print(f"  - Notices: {notice_count}")
    print(f"  - Exemplaires: {exemplaire_count}")
    print(f"  - Authors: {auteur_count}")


def main():
    print("=" * 80)
    print("ETL LOAD (FULL REFRESH): buf_clean.csv + bua_clean.csv -> MySQL")
    print("=" * 80)

    conn = get_connection()
    if not conn:
        return

    cursor = conn.cursor()

    try:
        buf, bua = load_two_csvs()
        full_refresh_load(conn, cursor, buf, bua)
        print_summary(cursor)
        print("\n✅ ETL LOAD COMPLETE")
    except Exception as e:
        print(f"\n❌ Error during load: {e}")
        conn.rollback()
    finally:
        cursor.close()
        conn.close()


if __name__ == "__main__":
    main()

