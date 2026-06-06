"""Temporary SIGB pipeline test runner with strict index normalization fixes."""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT_DIR = Path(__file__).resolve().parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

SCHEMA_PATH = ROOT_DIR / "sql" / "00_migration_clean.sql"
BUA_PATH = ROOT_DIR / "data" / "bua.xls"
BUF_PATH = ROOT_DIR / "data" / "buf.csv"
SKIPPED_PATH = ROOT_DIR / "skiped.csv"

TABLE_ORDER = [
    "langue",
    "editeur",
    "auteur",
    "classification",
    "matiere",
    "notice",
    "notice_auteur",
    "notice_matiere",
    "exemplaire",
]


def log(message: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")


def read_sql_statements(path: Path) -> list[str]:
    raw_text = path.read_text(encoding="utf-8")
    lines = [line for line in raw_text.splitlines() if not line.strip().startswith("--")]
    cleaned = "\n".join(lines)
    return [stmt.strip() for stmt in cleaned.split(";") if stmt.strip()]


def run_migration_sql(connection) -> None:
    log(f"Reading migration SQL: {SCHEMA_PATH}")
    statements = read_sql_statements(SCHEMA_PATH)
    log(f"Found {len(statements)} SQL statements")

    cursor = connection.cursor()
    try:
        for index, statement in enumerate(statements, start=1):
            preview = " ".join(statement.split())[:100]
            log(f"SQL {index}/{len(statements)}: {preview}")
            cursor.execute(statement)
        connection.commit()
        log("Migration SQL executed successfully")
    finally:
        cursor.close()


def read_source_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    from scripts.export import export_data_csv, export_data_excel

    if not BUA_PATH.exists():
        raise FileNotFoundError(f"Missing file: {BUA_PATH}")
    if not BUF_PATH.exists():
        raise FileNotFoundError(f"Missing file: {BUF_PATH}")

    log(f"Reading raw BUA source: {BUA_PATH}")
    bua_df = export_data_excel(BUA_PATH, sheet_name=0)
    log(f"BUA source rows={len(bua_df)} columns={list(bua_df.columns)}")

    log(f"Reading raw BUF source: {BUF_PATH}")
    buf_df = export_data_csv(BUF_PATH, encoding="latin1", sep=";")
    log(f"BUF source rows={len(buf_df)} columns={list(buf_df.columns)}")

    return bua_df, buf_df


def transform_sources(bua_df: pd.DataFrame, buf_df: pd.DataFrame) -> pd.DataFrame:
    from scripts.transform import transform

    log("Transforming BUA with scripts.transform.transform(lang_code='ara')")
    bua_t = transform(bua_df.copy(), "ara")
    log(f"BUA transformed rows={len(bua_t)} columns={list(bua_t.columns)}")

    log("Transforming BUF with scripts.transform.transform(lang_code='fre')")
    buf_t = transform(buf_df.copy(), "fre")
    log(f"BUF transformed rows={len(buf_t)} columns={list(buf_t.columns)}")

    combined = pd.concat([buf_t, bua_t], ignore_index=True)
    log(f"Combined transformed rows={len(combined)} columns={list(combined.columns)}")
    return combined


def build_etl_tables(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    from scripts.load import export_for_sql

    log("Building SQL-ready DataFrames with scripts.load.export_for_sql")
    return export_for_sql(df)


def ensure_connection():
    from scripts.db_connect import get_connection
    conn = get_connection()
    return conn


def safe_int(value) -> int | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        return int(float(text))
    except ValueError:
        return None


def append_skipped(skipped_rows: list[dict], entity: str, reason: str, row_data: dict) -> None:
    payload = {key: (None if pd.isna(value) else value) for key, value in row_data.items()}
    payload["entity"] = entity
    payload["reason"] = reason
    skipped_rows.append(payload)


def insert_lookup_values(cursor, connection, table: str, column: str, values: pd.Series) -> dict[str, int]:
    distinct = (
        values.dropna()
        .astype(str)
        .map(str.strip)
        .loc[lambda series: series != ""]
        .drop_duplicates()
    )

    log(f"Loading lookup table {table}: {len(distinct)} distinct values")
    for value in distinct:
        cursor.execute(f"INSERT IGNORE INTO {table} ({column}) VALUES (%s)", (value,))
    connection.commit()

    cursor.execute(f"SELECT * FROM {table}")
    rows = cursor.fetchall()
    # FIXED: String stripping handles text variations gracefully inside dictionary check keys
    return {str(row[1]).strip(): int(row[0]) for row in rows}


def load_to_mysql(connection, tables: dict[str, pd.DataFrame]) -> None:
    cursor = connection.cursor()
    skipped_rows: list[dict] = []
    try:
        log("Loading lookup tables")
        cursor.execute(
            "INSERT INTO langue (code_langue, libelle) VALUES (%s, %s) "
            "ON DUPLICATE KEY UPDATE libelle=VALUES(libelle)",
            ("fre", "Francais"),
        )
        cursor.execute(
            "INSERT INTO langue (code_langue, libelle) VALUES (%s, %s) "
            "ON DUPLICATE KEY UPDATE libelle=VALUES(libelle)",
            ("ara", "Arabe"),
        )
        connection.commit()

        cursor.execute("SELECT id_langue, code_langue FROM langue")
        langue_map = {str(row[1]).strip(): int(row[0]) for row in cursor.fetchall()}

        editeur_map = insert_lookup_values(cursor, connection, "editeur", "nom_editeur", tables["editeur"]["nom_editeur"])
        auteur_map = insert_lookup_values(cursor, connection, "auteur", "nom_complet", tables["auteur"]["nom_complet"])
        classification_map = insert_lookup_values(cursor, connection, "classification", "cote", tables["classification"]["cote"])
        matiere_map = insert_lookup_values(cursor, connection, "matiere", "libelle", tables["matiere"]["libelle"])

        notice_key_to_id: dict[str, int] = {}
        notice_inserted = 0
        notice_skipped_missing_data = 0
        notice_skipped_lookup = 0
        
        notice_rows = tables["notice"]
        log(f"Loading notices: {len(notice_rows)} rows")
        for index, row in notice_rows.iterrows():
            titre = str(row.get("Titre") or "").strip()
            notice_key = str(row.get("notice_key") or "").strip()
            nom_editeur = str(row.get("nom_editeur") or "").strip()
            code_langue = str(row.get("code_langue") or "").strip()
            cote = str(row.get("cote") or "").strip()

            if not titre or not notice_key:
                notice_skipped_missing_data += 1
                append_skipped(skipped_rows, "notice", "missing_titre_or_notice_key", row.to_dict())
                continue

            id_editeur = editeur_map.get(nom_editeur)
            id_classification = classification_map.get(cote)
            id_langue = langue_map.get(code_langue)

            if id_editeur is None or id_classification is None or id_langue is None:
                notice_skipped_lookup += 1
                append_skipped(skipped_rows, "notice", "missing_lookup_reference", row.to_dict())
                continue

            cursor.execute(
                """
                INSERT INTO notice
                    (titre, annee_pub, nb_pages, lieu_edition, id_editeur, id_langue, id_classification, note)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    titre,
                    safe_int(row.get("annee_pub")),
                    safe_int(row.get("nb_pages")),
                    (str(row.get("lieu_edition") or "").strip() or None),
                    id_editeur,
                    id_langue,
                    id_classification,
                    None,
                ),
            )
            notice_key_to_id[notice_key] = int(cursor.lastrowid)
            notice_inserted += 1

        connection.commit()
        log(f"Notice load complete: inserted={notice_inserted} unique_keys={len(notice_key_to_id)}")

        # --- Bridge tables insertion logic ---
        log(f"Loading notice_author bridge: {len(tables['notice_author'])} rows")
        for row in tables["notice_author"].itertuples(index=False):
            notice_id = notice_key_to_id.get(str(row.notice_key).strip())
            auteur_id = auteur_map.get(str(row.nom_complet or "").strip())
            if notice_id and auteur_id:
                cursor.execute(
                    "INSERT IGNORE INTO notice_auteur (id_notice, id_auteur) VALUES (%s, %s)",
                    (notice_id, auteur_id),
                )
        connection.commit()

        log(f"Loading notice_matiere bridge: {len(tables['notice_matiere'])} rows")
        for row in tables["notice_matiere"].itertuples(index=False):
            notice_id = notice_key_to_id.get(str(row.notice_key).strip())
            matiere_id = matiere_map.get(str(row.libelle or "").strip())
            if notice_id and matiere_id:
                cursor.execute(
                    "INSERT IGNORE INTO notice_matiere (id_notice, id_matiere) VALUES (%s, %s)",
                    (notice_id, matiere_id),
                )
        connection.commit()

        log(f"Loading exemplaire: {len(tables['exemplaire'])} rows")
        for row in tables["exemplaire"].itertuples(index=False):
            notice_id = notice_key_to_id.get(str(row.notice_key).strip())
            inv = str(row.num_inventaire or "").strip()
            # Clean floats like "29165.0" into proper tracking codes "29165"
            if inv.endswith(".0"):
                inv = inv[:-2]
                
            if notice_id and inv and inv != "nan":
                cursor.execute(
                    "INSERT IGNORE INTO exemplaire (num_inventaire, id_notice, etat) VALUES (%s, %s, %s)",
                    (inv, notice_id, "Disponible"),
                )
            else:
                append_skipped(skipped_rows, "exemplaire", "missing_notice_or_inventory", {
                    "notice_key": row.notice_key,
                    "num_inventaire": row.num_inventaire
                })
        connection.commit()

        skipped_df = pd.DataFrame(skipped_rows)
        skipped_df.to_csv(SKIPPED_PATH, index=False, encoding="utf-8-sig")
        log(f"Finished pipeline parsing script execution loop. Logged to {SKIPPED_PATH}.")

        for table in TABLE_ORDER:
            cursor.execute(f"SELECT COUNT(*) FROM {table}")
            count = int(cursor.fetchone()[0])
            log(f"Final count {table}: {count}")
    finally:
        cursor.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="SIGB temporary pipeline test")
    parser.add_argument("--reset", action="store_true", help="Run migration SQL setup step")
    args = parser.parse_args()

    connection = None
    try:
        connection = ensure_connection()
        if args.reset:
            run_migration_sql(connection)
        
        bua_df, buf_df = read_source_data()
        combined_df = transform_sources(bua_df, buf_df)
        tables = build_etl_tables(combined_df)
        load_to_mysql(connection, tables)
        return 0
    except Exception as exc:
        log(f"Pipeline test failed: {exc}")
        if connection is not None:
            connection.rollback()
        return 1
    finally:
        if connection is not None:
            connection.close()


if __name__ == "__main__":
    raise SystemExit(main())