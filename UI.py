"""
╔══════════════════════════════════════════════════════════════════════════════╗
║        SIGB — Système Intégré de Gestion Bibliothécaire                     ║
║        Interface Python/Tkinter  ←→  MySQL                                 ║
║                                                                              ║
║  PRÉREQUIS (installer une seule fois) :                                      ║
║    pip install mysql-connector-python python-dotenv pandas xlrd openpyxl    ║
║                                                                              ║
║  CONFIGURATION :                                                             ║
║    1. Configurer DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASSWORD dans .env  ║
║    2. Verifier que MySQL est accessible depuis ce poste                      ║
║    3. Lancer depuis la racine du projet : python UI.py                       ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import os
import csv
import re
import unicodedata
import threading
from datetime import datetime
from pathlib import Path

# ── ETL scripts (extract / transform / load)
from scripts import transform, load, export

# ── optional .env support
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parent / ".env")
except ImportError:
    pass

# ── optional MySQL driver
try:
    import mysql.connector
    MYSQL_OK = True
except ImportError:
    mysql = None
    MYSQL_OK = False

# ── optional pandas (needed for ETL import)
try:
    import pandas as pd
    PANDAS_OK = True
except ImportError:
    pd = None
    PANDAS_OK = False

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ══════════════════════════════════════════════════════════════════════════════
#  PATHS
# ══════════════════════════════════════════════════════════════════════════════

ROOT_DIR    = Path(__file__).resolve().parent
SCHEMA_PATH = ROOT_DIR / "sql" / "00_migration_clean.sql"

# ══════════════════════════════════════════════════════════════════════════════
#  DATABASE CREDENTIALS  (overridden by .env)
# ══════════════════════════════════════════════════════════════════════════════

DB_HOST     = os.getenv("DB_HOST", "localhost")
DB_PORT     = int(os.getenv("DB_PORT", "3306"))
DB_NAME     = os.getenv("DB_NAME", "sigb")
DB_USER     = os.getenv("DB_USER", "sigb_user")
DB_PASSWORD = os.getenv("DB_PASSWORD", "")

# ══════════════════════════════════════════════════════════════════════════════
#  THEME & CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════

APP_TITLE   = "SIGB — MySQL"
APP_VERSION = "2.0"

C_BG        = "#F0F2F5"
C_SURFACE   = "#FFFFFF"
C_PRIMARY   = "#1A3A5C"
C_PRIMARY_L = "#2E6DA4"
C_ACCENT    = "#C0392B"
C_GOLD      = "#D4A017"
C_TEXT      = "#1A1A2E"
C_MUTED     = "#6B7A99"
C_BORDER    = "#DDE3EC"
C_ROW_ODD   = "#F7F9FC"
C_ROW_EVEN  = "#FFFFFF"
C_SEL       = "#D0E4F7"
C_SUCCESS   = "#2A6E1A"
C_ONLINE    = "#27AE60"
C_ACCENT_R  = "#E74C3C"

FONT_TITLE  = ("Segoe UI", 13, "bold")
FONT_HEAD   = ("Segoe UI", 10, "bold")
FONT_BODY   = ("Segoe UI", 9)
FONT_SMALL  = ("Segoe UI", 8)

# Columns displayed in the main grid (mapped to SQL aliases in search query)
COLUMNS = [
    "N_INVENTAIRE", "COTE",     "TITRE",
    "NOM_AUTEUR",   "NOM_LIEU", "NOM_EDITEUR",
    "ANNEE",        "NB_PAGES", "NOM_MATIERE",
    "CODE_FONDS",
]
COL_LABELS = {
    "N_INVENTAIRE": "N° Inv.",
    "COTE":         "Cote",
    "TITRE":        "Titre",
    "NOM_AUTEUR":   "Auteur",
    "NOM_LIEU":     "Lieu",
    "NOM_EDITEUR":  "Éditeur",
    "ANNEE":        "Année",
    "NB_PAGES":     "Pages",
    "NOM_MATIERE":  "Matière",
    "CODE_FONDS":   "Fonds",
}
COL_WIDTHS = {
    "N_INVENTAIRE": 80,
    "COTE":         110,
    "TITRE":        320,
    "NOM_AUTEUR":   175,
    "NOM_LIEU":     95,
    "NOM_EDITEUR":  125,
    "ANNEE":        60,
    "NB_PAGES":     60,
    "NOM_MATIERE":  200,
    "CODE_FONDS":   60,
}


# ══════════════════════════════════════════════════════════════════════════════
#  DATA ACCESS LAYER  (MySQLDAO)
#  Schema: sql/00_migration_clean.sql
#  Tables: langue, editeur, auteur, classification, matiere,
#          notice, notice_auteur, notice_matiere, exemplaire
# ══════════════════════════════════════════════════════════════════════════════

class MySQLDAO:
    """All MySQL interactions go through this class."""

    # Maps UI fonds codes to DB language codes and vice-versa
    LANG_BY_FONDS = {"BUA": "ara", "BUF": "fre"}
    FONDS_BY_LANG = {"ara": "BUA", "fre": "BUF"}

    def __init__(self, host=DB_HOST, port=DB_PORT, database=DB_NAME,
                 user=DB_USER, password=DB_PASSWORD):
        self.host     = host
        self.port     = int(port)
        self.database = database
        self.user     = user
        self.password = password
        self.conn     = None
        self._connected = False

    # ── Connection management ─────────────────────────────────────────────────

    def connect(self):
        if not MYSQL_OK:
            raise RuntimeError(
                "mysql-connector-python non installé.\n"
                "Exécutez : pip install mysql-connector-python python-dotenv"
            )
        self.conn = mysql.connector.connect(
            host=self.host, port=self.port, database=self.database,
            user=self.user, password=self.password,
            charset="utf8mb4", connection_timeout=10, use_pure=True,
        )
        if not self.conn.is_connected():
            raise RuntimeError("Connexion créée mais inactive.")
        self._connected = True
        return True

    def disconnect(self):
        if self.conn:
            try:
                self.conn.close()
            except Exception:
                pass
        self._connected = False
        self.conn = None

    def is_connected(self):
        return self._connected and self.conn is not None

    def ping(self):
        try:
            if self.conn:
                self.conn.ping(reconnect=True, attempts=1, delay=0)
                return True
        except Exception:
            self._connected = False
        return False

    # ── Schema management ─────────────────────────────────────────────────────

    def schema_exists(self):
        """Returns True if the 'notice' table already exists in the database."""
        try:
            cur = self.conn.cursor()
            cur.execute("SHOW TABLES LIKE 'notice'")
            res = cur.fetchone()
            cur.close()
            return res is not None
        except Exception:
            return False

    def init_schema(self):
        """
        Executes sql/00_migration_clean.sql to create all tables.
        Called only when schema_exists() is False.
        """
        if not SCHEMA_PATH.exists():
            raise FileNotFoundError(f"Schema SQL introuvable : {SCHEMA_PATH}")

        sql_script = SCHEMA_PATH.read_text(encoding="utf-8")

        # Strip comment lines then split on semicolons
        lines = [l for l in sql_script.splitlines() if not l.strip().startswith("--")]
        statements = [s.strip() for s in "\n".join(lines).split(";") if s.strip()]

        cur = self.conn.cursor()
        try:
            for stmt in statements:
                cur.execute(stmt)
            self.conn.commit()
        finally:
            cur.close()
        return True

    # ── Search / read ─────────────────────────────────────────────────────────

    def search(self, keyword="", field="TOUS", fonds="TOUS",
               yr_from=None, yr_to=None,
               sort_col="N_INVENTAIRE", sort_asc=True,
               limit=100, offset=0):
        """
        Paginated search across the normalised MySQL schema.
        Returns (rows: list[dict], total: int).
        Columns come directly from base tables to avoid subquery overhead.
        """
        field_map = {
            "TOUS":    None,
            "TITRE":   "n.titre",
            "AUTEUR":  "a.nom_complet",
            "COTE":    "c.cote",
            "MATIERE": "m.libelle",
            "EDITEUR": "ed.nom_editeur",
            "LIEU":    "n.lieu_edition",
        }
        sql_field = field_map.get(field.upper())

        conditions, params = [], []

        if keyword:
            kw = f"%{keyword.upper()}%"
            if sql_field:
                conditions.append(f"UPPER(COALESCE({sql_field}, '')) LIKE %s")
                params.append(kw)
            else:
                conditions.append(
                    "(UPPER(COALESCE(n.titre, '')) LIKE %s "
                    "OR UPPER(COALESCE(a.nom_complet, '')) LIKE %s "
                    "OR UPPER(COALESCE(c.cote, '')) LIKE %s "
                    "OR UPPER(COALESCE(m.libelle, '')) LIKE %s "
                    "OR UPPER(COALESCE(ed.nom_editeur, '')) LIKE %s)"
                )
                params.extend([kw] * 5)

        if fonds and fonds != "TOUS":
            conditions.append("l.code_langue = %s")
            params.append("ara" if fonds == "BUA" else "fre")

        if yr_from:
            conditions.append("n.annee_pub >= %s")
            params.append(int(yr_from))
        if yr_to:
            conditions.append("n.annee_pub <= %s")
            params.append(int(yr_to))

        where = ("WHERE " + " AND ".join(conditions)) if conditions else ""

        direction = "ASC" if sort_asc else "DESC"
        sort_map = {
            "N_INVENTAIRE": "ex.num_inventaire",
            "COTE":         "c.cote",
            "TITRE":        "n.titre",
            "NOM_AUTEUR":   "a.nom_complet",
            "NOM_LIEU":     "n.lieu_edition",
            "NOM_EDITEUR":  "ed.nom_editeur",
            "ANNEE":        "n.annee_pub",
            "NB_PAGES":     "n.nb_pages",
            "NOM_MATIERE":  "m.libelle",
            "CODE_FONDS":   "l.code_langue",
        }
        safe_sort = sort_map.get(sort_col, "ex.num_inventaire")

        joins = """
            FROM notice n
            LEFT JOIN langue         l  ON l.id_langue        = n.id_langue
            LEFT JOIN classification c  ON c.id_classification= n.id_classification
            LEFT JOIN editeur        ed ON ed.id_editeur       = n.id_editeur
            LEFT JOIN exemplaire     ex ON ex.id_notice        = n.id_notice
            LEFT JOIN notice_auteur  na ON na.id_notice        = n.id_notice
            LEFT JOIN auteur         a  ON a.id_auteur         = na.id_auteur
            LEFT JOIN notice_matiere nm ON nm.id_notice        = n.id_notice
            LEFT JOIN matiere        m  ON m.id_matiere        = nm.id_matiere
        """

        sql_data = f"""
            SELECT
                n.id_notice          AS ID_NOTICE,
                ex.num_inventaire    AS N_INVENTAIRE,
                c.cote               AS COTE,
                n.titre              AS TITRE,
                GROUP_CONCAT(DISTINCT a.nom_complet  SEPARATOR '; ') AS NOM_AUTEUR,
                n.lieu_edition       AS NOM_LIEU,
                ed.nom_editeur       AS NOM_EDITEUR,
                n.annee_pub          AS ANNEE,
                n.nb_pages           AS NB_PAGES,
                GROUP_CONCAT(DISTINCT m.libelle      SEPARATOR '; ') AS NOM_MATIERE,
                CASE l.code_langue
                    WHEN 'ara' THEN 'BUA'
                    WHEN 'fre' THEN 'BUF'
                    ELSE UPPER(l.code_langue)
                END                  AS CODE_FONDS,
                n.date_catalogage    AS DATE_AJOUT
            {joins}
            {where}
            GROUP BY ex.id_exemplaire, n.id_notice
            ORDER BY ({safe_sort} IS NULL), {safe_sort} {direction}
            LIMIT {int(limit)} OFFSET {int(offset)}
        """

        sql_count = f"""
            SELECT COUNT(DISTINCT ex.id_exemplaire)
            {joins}
            {where}
        """

        conn2 = mysql.connector.connect(
            host=self.host, port=self.port, database=self.database,
            user=self.user, password=self.password,
            charset="utf8mb4", connection_timeout=10, use_pure=True,
        )
        try:
            cur = conn2.cursor()
            cur.execute(sql_count, tuple(params))
            total = cur.fetchone()[0]
            cur.close()

            cur = conn2.cursor(dictionary=True)
            cur.execute(sql_data, tuple(params))
            rows = cur.fetchall()
            cur.close()
        finally:
            try:
                conn2.close()
            except Exception:
                pass
        return rows, total

    def get_notice(self, id_notice):
        """Fetch a single notice by ID for editing (uses _base_select)."""
        cur = self.conn.cursor(dictionary=True)
        cur.execute(f"""
            SELECT *
            FROM ({self._base_select()}) v
            WHERE ID_NOTICE = %s
            LIMIT 1
        """, (id_notice,))
        row = cur.fetchone()
        cur.close()
        return row

    def _base_select(self):
        """Full aggregated SELECT used for single-record retrieval (get_notice)."""
        return """
            SELECT
                n.id_notice          AS ID_NOTICE,
                ex.num_inventaire    AS N_INVENTAIRE,
                c.cote               AS COTE,
                n.titre              AS TITRE,
                GROUP_CONCAT(DISTINCT a.nom_complet  SEPARATOR '; ') AS NOM_AUTEUR,
                n.lieu_edition       AS NOM_LIEU,
                ed.nom_editeur       AS NOM_EDITEUR,
                n.annee_pub          AS ANNEE,
                n.nb_pages           AS NB_PAGES,
                GROUP_CONCAT(DISTINCT m.libelle      SEPARATOR '; ') AS NOM_MATIERE,
                CASE l.code_langue
                    WHEN 'ara' THEN 'BUA'
                    WHEN 'fre' THEN 'BUF'
                    ELSE UPPER(l.code_langue)
                END                  AS CODE_FONDS,
                n.date_catalogage    AS DATE_AJOUT
            FROM notice n
            LEFT JOIN langue         l  ON l.id_langue         = n.id_langue
            LEFT JOIN classification c  ON c.id_classification = n.id_classification
            LEFT JOIN editeur        ed ON ed.id_editeur        = n.id_editeur
            LEFT JOIN exemplaire     ex ON ex.id_notice         = n.id_notice
            LEFT JOIN notice_auteur  na ON na.id_notice         = n.id_notice
            LEFT JOIN auteur         a  ON a.id_auteur          = na.id_auteur
            LEFT JOIN notice_matiere nm ON nm.id_notice         = n.id_notice
            LEFT JOIN matiere        m  ON m.id_matiere         = nm.id_matiere
            GROUP BY ex.id_exemplaire, n.id_notice
        """

    def get_stats(self):
        """Returns summary counts for the status bar."""
        conn2 = mysql.connector.connect(
            host=self.host, port=self.port, database=self.database,
            user=self.user, password=self.password,
            charset="utf8mb4", connection_timeout=10, use_pure=True,
        )
        stats = {}
        try:
            cur = conn2.cursor()
            cur.execute("SELECT COUNT(*) FROM notice")
            stats["total"] = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM notice n JOIN langue l ON l.id_langue=n.id_langue WHERE l.code_langue='ara'")
            stats["bua"] = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM notice n JOIN langue l ON l.id_langue=n.id_langue WHERE l.code_langue='fre'")
            stats["buf"] = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM matiere")
            stats["matieres"] = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM auteur")
            stats["auteurs"] = cur.fetchone()[0]
            cur.close()
        finally:
            try:
                conn2.close()
            except Exception:
                pass
        return stats

    # ── Single-record CRUD ────────────────────────────────────────────────────
    # These methods are used by ItemDialog (add/edit/delete individual notices)

    def _clean(self, value):
        if value is None:
            return ""
        text = str(value).strip()
        return "".join(c for c in text if unicodedata.category(c)[0] != "C").strip()

    def _safe_int(self, value):
        v = self._clean(value)
        try:
            return int(float(v)) if v else None
        except (TypeError, ValueError):
            return None

    def _split_values(self, value):
        v = self._clean(value)
        if not v:
            return []
        return [p.strip() for p in re.split(r"\s*[;/]\s*", v) if p.strip()]

    def _get_langue_id(self, fonds_code):
        code = self.LANG_BY_FONDS.get((fonds_code or "BUF").upper(), "fre")
        cur = self.conn.cursor()
        cur.execute("SELECT id_langue FROM langue WHERE code_langue=%s", (code,))
        row = cur.fetchone()
        if row:
            cur.close()
            return row[0]
        libelle = {"ara": "Arabe", "fre": "Francais", "eng": "Anglais"}.get(code, code)
        cur.execute("INSERT INTO langue (code_langue, libelle) VALUES (%s, %s)", (code, libelle))
        self.conn.commit()
        new_id = cur.lastrowid
        cur.close()
        return new_id

    def _get_or_create(self, table, key_col, id_col, value):
        """Generic get-or-create for lookup tables (editeur, classification, matiere, auteur)."""
        value = self._clean(value)
        if not value:
            return None
        cur = self.conn.cursor()
        cur.execute(f"INSERT IGNORE INTO {table} ({key_col}) VALUES (%s)", (value,))
        cur.execute(f"SELECT {id_col} FROM {table} WHERE {key_col}=%s", (value,))
        row = cur.fetchone()
        cur.close()
        return row[0] if row else None

    def _link_authors_and_subjects(self, cur, id_notice, auteurs, matieres):
        for auteur in self._split_values(auteurs):
            id_a = self._get_or_create("auteur", "nom_complet", "id_auteur", auteur)
            if id_a:
                cur.execute(
                    "INSERT IGNORE INTO notice_auteur (id_notice, id_auteur) VALUES (%s, %s)",
                    (id_notice, id_a),
                )
        for matiere in self._split_values(matieres):
            id_m = self._get_or_create("matiere", "libelle", "id_matiere", matiere)
            if id_m:
                cur.execute(
                    "INSERT IGNORE INTO notice_matiere (id_notice, id_matiere) VALUES (%s, %s)",
                    (id_notice, id_m),
                )

    def _upsert_exemplaire(self, cur, id_notice, inventaire):
        inv = self._clean(inventaire)
        if not inv:
            return
        cur.execute("SELECT id_exemplaire FROM exemplaire WHERE id_notice=%s LIMIT 1", (id_notice,))
        row = cur.fetchone()
        if row:
            cur.execute("UPDATE exemplaire SET num_inventaire=%s WHERE id_exemplaire=%s", (inv, row[0]))
        else:
            cur.execute("INSERT INTO exemplaire (num_inventaire, id_notice) VALUES (%s, %s)", (inv, id_notice))

    def add_notice(self, data):
        inv = self._clean(data.get("N_INVENTAIRE"))
        # Check uniqueness
        cur = self.conn.cursor()
        cur.execute("SELECT 1 FROM exemplaire WHERE num_inventaire=%s LIMIT 1", (inv,))
        if cur.fetchone():
            cur.close()
            raise RuntimeError(f"Le numéro d'inventaire existe déjà : {inv}")
        cur.execute("""
            INSERT INTO notice (titre, annee_pub, nb_pages, lieu_edition,
                                id_editeur, id_langue, id_classification)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """, (
            self._clean(data.get("TITRE")),
            self._safe_int(data.get("ANNEE")),
            self._safe_int(data.get("NB_PAGES")),
            self._clean(data.get("NOM_LIEU")) or None,
            self._get_or_create("editeur", "nom_editeur", "id_editeur", data.get("NOM_EDITEUR")),
            self._get_langue_id(data.get("CODE_FONDS", "BUF")),
            self._get_or_create("classification", "cote", "id_classification", data.get("COTE")),
        ))
        id_notice = cur.lastrowid
        self._link_authors_and_subjects(cur, id_notice, data.get("NOM_AUTEUR"), data.get("NOM_MATIERE"))
        self._upsert_exemplaire(cur, id_notice, inv)
        self.conn.commit()
        cur.close()

    def update_notice(self, id_notice, data):
        cur = self.conn.cursor()
        cur.execute("""
            UPDATE notice
            SET titre=%s, annee_pub=%s, nb_pages=%s, lieu_edition=%s,
                id_editeur=%s, id_langue=%s, id_classification=%s
            WHERE id_notice=%s
        """, (
            self._clean(data.get("TITRE")),
            self._safe_int(data.get("ANNEE")),
            self._safe_int(data.get("NB_PAGES")),
            self._clean(data.get("NOM_LIEU")) or None,
            self._get_or_create("editeur", "nom_editeur", "id_editeur", data.get("NOM_EDITEUR")),
            self._get_langue_id(data.get("CODE_FONDS", "BUF")),
            self._get_or_create("classification", "cote", "id_classification", data.get("COTE")),
            id_notice,
        ))
        cur.execute("DELETE FROM notice_auteur  WHERE id_notice=%s", (id_notice,))
        cur.execute("DELETE FROM notice_matiere WHERE id_notice=%s", (id_notice,))
        self._link_authors_and_subjects(cur, id_notice, data.get("NOM_AUTEUR"), data.get("NOM_MATIERE"))
        self._upsert_exemplaire(cur, id_notice, data.get("N_INVENTAIRE"))
        self.conn.commit()
        cur.close()

    def delete_notice(self, id_notice):
        cur = self.conn.cursor()
        cur.execute("DELETE FROM exemplaire     WHERE id_notice=%s", (id_notice,))
        cur.execute("DELETE FROM notice         WHERE id_notice=%s", (id_notice,))
        self.conn.commit()
        cur.close()

    # ── Bulk ETL import ───────────────────────────────────────────────────────

    def load_etl_tables(self, tables, progress_cb=None):
        """
        Load ETL-produced DataFrames into the database.
        tables: dict returned by scripts.load.export_for_sql()
        Keys expected: langue, editeur, auteur, classification, matiere,
                       notice, notice_author, notice_matiere, exemplaire
        """

        def _safe_int(value):
            if pd.isna(value):
                return None
            text = str(value).strip()
            try:
                return int(float(text)) if text else None
            except ValueError:
                return None

        def _insert_lookup(cursor, table, column, values):
            """Insert distinct values into a lookup table, return {value: id} map."""
            distinct = (
                values.dropna().astype(str).map(str.strip)
                .loc[lambda s: s != ""].drop_duplicates()
            )
            for v in distinct:
                cursor.execute(f"INSERT IGNORE INTO {table} ({column}) VALUES (%s)", (v,))
            self.conn.commit()
            cursor.execute(f"SELECT * FROM {table}")
            return {str(r[1]).strip(): int(r[0]) for r in cursor.fetchall()}

        cur = self.conn.cursor()
        try:
            if progress_cb:
                progress_cb("Chargement des tables de référence...", 10)

            # Ensure both language rows exist
            for code, libelle in [("fre", "Francais"), ("ara", "Arabe")]:
                cur.execute(
                    "INSERT INTO langue (code_langue, libelle) VALUES (%s,%s) "
                    "ON DUPLICATE KEY UPDATE libelle=VALUES(libelle)",
                    (code, libelle),
                )
            self.conn.commit()

            cur.execute("SELECT id_langue, code_langue FROM langue")
            langue_map = {str(r[1]).strip(): int(r[0]) for r in cur.fetchall()}

            editeur_map        = _insert_lookup(cur, "editeur",        "nom_editeur", tables["editeur"]["nom_editeur"])
            auteur_map         = _insert_lookup(cur, "auteur",         "nom_complet", tables["auteur"]["nom_complet"])
            classification_map = _insert_lookup(cur, "classification", "cote",        tables["classification"]["cote"])
            matiere_map        = _insert_lookup(cur, "matiere",        "libelle",     tables["matiere"]["libelle"])

            if progress_cb:
                progress_cb("Chargement des notices...", 40)

            notice_key_to_id = {}
            notice_inserted  = 0

            for _, row in tables["notice"].iterrows():
                titre       = str(row.get("Titre")       or "").strip()
                notice_key  = str(row.get("notice_key")  or "").strip()
                nom_editeur = str(row.get("nom_editeur") or "").strip()
                code_langue = str(row.get("code_langue") or "").strip()
                cote        = str(row.get("cote")        or "").strip()

                if not titre or not notice_key:
                    continue

                id_editeur        = editeur_map.get(nom_editeur)
                id_classification = classification_map.get(cote)
                id_langue         = langue_map.get(code_langue)

                if None in (id_editeur, id_classification, id_langue):
                    continue

                cur.execute("""
                    INSERT INTO notice
                        (titre, annee_pub, nb_pages, lieu_edition,
                         id_editeur, id_langue, id_classification, note)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    titre,
                    _safe_int(row.get("annee_pub")),
                    _safe_int(row.get("nb_pages")),
                    str(row.get("lieu_edition") or "").strip() or None,
                    id_editeur, id_langue, id_classification,
                    None,
                ))
                notice_key_to_id[notice_key] = int(cur.lastrowid)
                notice_inserted += 1

            self.conn.commit()

            if progress_cb:
                progress_cb("Chargement des auteurs & matières...", 70)

            for row in tables["notice_author"].itertuples(index=False):
                nid = notice_key_to_id.get(str(row.notice_key).strip())
                aid = auteur_map.get(str(row.nom_complet or "").strip())
                if nid and aid:
                    cur.execute(
                        "INSERT IGNORE INTO notice_auteur (id_notice, id_auteur) VALUES (%s,%s)",
                        (nid, aid),
                    )
            self.conn.commit()

            for row in tables["notice_matiere"].itertuples(index=False):
                nid = notice_key_to_id.get(str(row.notice_key).strip())
                mid = matiere_map.get(str(row.libelle or "").strip())
                if nid and mid:
                    cur.execute(
                        "INSERT IGNORE INTO notice_matiere (id_notice, id_matiere) VALUES (%s,%s)",
                        (nid, mid),
                    )
            self.conn.commit()

            if progress_cb:
                progress_cb("Chargement des exemplaires...", 90)

            for row in tables["exemplaire"].itertuples(index=False):
                nid = notice_key_to_id.get(str(row.notice_key).strip())
                inv = str(row.num_inventaire or "").strip()
                if inv.endswith(".0"):
                    inv = inv[:-2]
                if nid and inv and inv.lower() != "nan":
                    cur.execute(
                        "INSERT IGNORE INTO exemplaire (num_inventaire, id_notice, etat) VALUES (%s,%s,%s)",
                        (inv, nid, "Disponible"),
                    )
            self.conn.commit()

            if progress_cb:
                progress_cb("Importation terminée !", 100)

            return notice_inserted
        finally:
            cur.close()


# ══════════════════════════════════════════════════════════════════════════════
#  CONNECTION DIALOG
# ══════════════════════════════════════════════════════════════════════════════

class ConnectDialog(tk.Toplevel):
    def __init__(self, parent, on_connect):
        super().__init__(parent)
        self.title("Connexion MySQL")
        self.resizable(False, False)
        self.configure(bg=C_BG)
        self.on_connect = on_connect
        self._build()
        self.grab_set()
        self.focus_force()

    def _build(self):
        hdr = tk.Frame(self, bg=C_PRIMARY, pady=14, padx=20)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🔐 Connexion MySQL",            font=FONT_TITLE, fg="white",  bg=C_PRIMARY).pack(anchor="w")
        tk.Label(hdr, text="Système d'authentification SIGB", font=FONT_SMALL, fg=C_BORDER, bg=C_PRIMARY).pack(anchor="w")

        body = tk.Frame(self, bg=C_BG, padx=20, pady=15)
        body.pack(fill="both", expand=True)

        fields = [
            ("Hôte de base de données :", "host", DB_HOST),
            ("Port réseau :",             "port", str(DB_PORT)),
            ("Nom du schéma (DB) :",      "db",   DB_NAME),
            ("Identifiant utilisateur :",  "user", DB_USER),
            ("Mot de passe :",             "pwd",  DB_PASSWORD),
        ]
        self.entries = {}
        for label, key, default in fields:
            row = tk.Frame(body, bg=C_BG, pady=4)
            row.pack(fill="x")
            tk.Label(row, text=label, font=FONT_BODY, bg=C_BG, width=22, anchor="w").pack(side="left")
            ent = tk.Entry(row, font=FONT_BODY, bg=C_SURFACE, highlightthickness=1,
                           highlightbackground=C_BORDER, highlightcolor=C_PRIMARY_L,
                           show="*" if key == "pwd" else "")
            ent.insert(0, default)
            ent.pack(side="right", fill="x", expand=True)
            self.entries[key] = ent

        btn_frame = tk.Frame(body, bg=C_BG, pady=10)
        btn_frame.pack(fill="x")
        self.btn_submit = tk.Button(
            btn_frame, text="Se Connecter", font=FONT_HEAD,
            bg=C_PRIMARY, fg="white", activebackground=C_PRIMARY_L,
            activeforeground="white", padx=15, pady=4, bd=0,
            cursor="hand2", command=self._validate,
        )
        self.btn_submit.pack(side="right")

    def _validate(self):
        h = self.entries["host"].get().strip()
        p = self.entries["port"].get().strip()
        d = self.entries["db"].get().strip()
        u = self.entries["user"].get().strip()
        w = self.entries["pwd"].get()

        if not all([h, p, d, u]):
            messagebox.showerror("Erreur", "Tous les champs sauf le mot de passe sont obligatoires.")
            return

        self.btn_submit.config(state="disabled", text="Connexion...")
        self.update()

        dao = MySQLDAO(host=h, port=p, database=d, user=u, password=w)
        try:
            dao.connect()
            self.on_connect(dao)
            self.destroy()
        except Exception as e:
            messagebox.showerror("Échec de connexion", f"Erreur rencontrée :\n{e}")
            self.btn_submit.config(state="normal", text="Se Connecter")


# ══════════════════════════════════════════════════════════════════════════════
#  NOTICE CREATE / EDIT DIALOG
# ══════════════════════════════════════════════════════════════════════════════

class ItemDialog(tk.Toplevel):
    def __init__(self, parent, dao, record=None, on_save=None):
        super().__init__(parent)
        self.dao      = dao
        self.record   = record
        self.on_save  = on_save
        self.title("Notice — Modification" if record else "Notice — Création")
        self.resizable(False, False)
        self.configure(bg=C_BG)
        self._build()
        self.grab_set()
        self.focus_force()

    def _build(self):
        hdr_bg  = C_PRIMARY if self.record else C_PRIMARY_L
        hdr_txt = "📝 Édition de la Notice" if self.record else "➕ Ajouter une Notice"
        hdr = tk.Frame(self, bg=hdr_bg, pady=12, padx=20)
        hdr.pack(fill="x")
        tk.Label(hdr, text=hdr_txt, font=FONT_TITLE, fg="white", bg=hdr_bg).pack(anchor="w")

        body = tk.Frame(self, bg=C_BG, padx=20, pady=15)
        body.pack(fill="both", expand=True)

        self.vars = {}
        fields = [
            ("N° Inventaire (*unique) :",   "N_INVENTAIRE"),
            ("Cote Classification :",        "COTE"),
            ("Titre Principal :",            "TITRE"),
            ("Auteur(s) [séparés par ';'] :","NOM_AUTEUR"),
            ("Matière(s) [séparés par ';']:","NOM_MATIERE"),
            ("Nom Éditeur :",                "NOM_EDITEUR"),
            ("Lieu / Ville Édition :",       "NOM_LIEU"),
            ("Année Publication :",          "ANNEE"),
            ("Nombre de Pages :",            "NB_PAGES"),
        ]
        for lbl_txt, col in fields:
            row = tk.Frame(body, bg=C_BG, pady=3)
            row.pack(fill="x")
            tk.Label(row, text=lbl_txt, font=FONT_BODY, bg=C_BG, width=25, anchor="w").pack(side="left")
            var = tk.StringVar()
            if self.record and col in self.record:
                val = self.record[col]
                var.set("" if val is None else str(val))
            ent = tk.Entry(row, textvariable=var, font=FONT_BODY, bg=C_SURFACE,
                           highlightthickness=1, highlightbackground=C_BORDER)
            ent.pack(side="right", fill="x", expand=True)
            self.vars[col] = var

        row_f = tk.Frame(body, bg=C_BG, pady=5)
        row_f.pack(fill="x")
        tk.Label(row_f, text="Fonds Documentaire :", font=FONT_BODY, bg=C_BG, width=25, anchor="w").pack(side="left")
        default_fonds = self.record.get("CODE_FONDS", "BUF") if self.record else "BUF"
        self.var_fonds = tk.StringVar(value=default_fonds)
        ttk.Combobox(row_f, textvariable=self.var_fonds, values=["BUA", "BUF"],
                     state="readonly", font=FONT_BODY).pack(side="right", fill="x", expand=True)

        btn_frame = tk.Frame(body, bg=C_BG, pady=10)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="Enregistrer", font=FONT_HEAD, bg=C_SUCCESS, fg="white",
                  padx=15, pady=4, bd=0, cursor="hand2", command=self._save).pack(side="right", padx=5)
        tk.Button(btn_frame, text="Annuler", font=FONT_BODY, bg=C_MUTED, fg="white",
                  padx=10, pady=4, bd=0, cursor="hand2", command=self.destroy).pack(side="right")

    def _save(self):
        data = {col: var.get().strip() for col, var in self.vars.items()}
        data["CODE_FONDS"] = self.var_fonds.get()

        if not data["N_INVENTAIRE"] or not data["TITRE"]:
            messagebox.showerror("Champs Requis", "N° Inventaire et Titre sont obligatoires.")
            return

        try:
            if self.record:
                self.dao.update_notice(self.record["ID_NOTICE"], data)
            else:
                self.dao.add_notice(data)
            if self.on_save:
                self.on_save()
            self.destroy()
        except Exception as e:
            messagebox.showerror("Erreur Sauvegarde", f"Action impossible :\n{e}")


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION WINDOW
# ══════════════════════════════════════════════════════════════════════════════

class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} v{APP_VERSION}")
        self.geometry("1300x750")
        self.minsize(1100, 650)
        self.configure(bg=C_BG)

        self.dao        = None
        self._page      = 0
        self._limit     = 100
        self._sort_col  = "N_INVENTAIRE"
        self._sort_asc  = True

        self._init_variables()
        self._style_treeview()
        self._build_layout()

        # Show login dialog shortly after the window opens
        self.after(100, self._prompt_login)

    # ── Variables ─────────────────────────────────────────────────────────────

    def _init_variables(self):
        self.var_kw    = tk.StringVar()
        self.var_field = tk.StringVar(value="TOUS")
        self.var_fonds = tk.StringVar(value="TOUS")
        self.var_yr1   = tk.StringVar()
        self.var_yr2   = tk.StringVar()
        self._status_var = tk.StringVar(value="Hors-ligne — Authentification requise.")
        self._stats_var  = tk.StringVar(value="Total : -- | BUA : -- | BUF : -- | Auteurs : -- | Matières : --")

    # ── Styling ───────────────────────────────────────────────────────────────

    def _style_treeview(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background=C_SURFACE, foreground=C_TEXT,
                        rowheight=24, fieldbackground=C_SURFACE, font=FONT_BODY)
        style.map("Treeview",
                  background=[("selected", C_SEL)],
                  foreground=[("selected", C_PRIMARY)])
        style.configure("Treeview.Heading", background=C_BORDER,
                        foreground=C_TEXT, font=FONT_HEAD, padding=5)

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_layout(self):
        # Top bar
        top = tk.Frame(self, bg=C_PRIMARY, pady=10, padx=15)
        top.pack(fill="x")
        tk.Label(top, text="📖 SIGB", font=FONT_TITLE, fg="white", bg=C_PRIMARY).pack(side="left")
        self.lbl_status_badge = tk.Label(top, text="DÉCONNECTÉ", font=FONT_SMALL,
                                          bg=C_ACCENT, fg="white", padx=6, pady=2)
        self.lbl_status_badge.pack(side="left", padx=12)

        # Search bar
        sf = tk.LabelFrame(self, text=" Moteur de Recherche Documentaire ",
                           font=FONT_HEAD, bg=C_SURFACE, fg=C_PRIMARY,
                           padx=15, pady=10, bd=1)
        sf.pack(fill="x", padx=15, pady=10)

        r1 = tk.Frame(sf, bg=C_SURFACE)
        r1.pack(fill="x", pady=2)
        tk.Label(r1, text="Expression recherchée :", font=FONT_BODY, bg=C_SURFACE).pack(side="left")
        ent_kw = tk.Entry(r1, textvariable=self.var_kw, font=FONT_BODY, bg=C_SURFACE,
                          width=35, highlightthickness=1, highlightbackground=C_BORDER)
        ent_kw.pack(side="left", padx=5)
        ent_kw.bind("<Return>", lambda e: self._search_reset_page())

        tk.Label(r1, text="Cible :", font=FONT_BODY, bg=C_SURFACE, padx=10).pack(side="left")
        ttk.Combobox(r1, textvariable=self.var_field, font=FONT_BODY, width=12, state="readonly",
                     values=["TOUS", "TITRE", "AUTEUR", "COTE", "MATIERE", "EDITEUR", "LIEU"]).pack(side="left")

        tk.Label(r1, text="Fonds :", font=FONT_BODY, bg=C_SURFACE, padx=10).pack(side="left")
        ttk.Combobox(r1, textvariable=self.var_fonds, font=FONT_BODY, width=8, state="readonly",
                     values=["TOUS", "BUA", "BUF"]).pack(side="left")

        r2 = tk.Frame(sf, bg=C_SURFACE, pady=4)
        r2.pack(fill="x")
        tk.Label(r2, text="Période (Années) entre :", font=FONT_BODY, bg=C_SURFACE).pack(side="left")
        tk.Entry(r2, textvariable=self.var_yr1, font=FONT_BODY, width=6,
                 highlightthickness=1, highlightbackground=C_BORDER).pack(side="left", padx=4)
        tk.Label(r2, text="et", font=FONT_BODY, bg=C_SURFACE).pack(side="left", padx=2)
        tk.Entry(r2, textvariable=self.var_yr2, font=FONT_BODY, width=6,
                 highlightthickness=1, highlightbackground=C_BORDER).pack(side="left", padx=4)
        tk.Button(r2, text="❌ Nettoyer", font=FONT_BODY, bg=C_MUTED, fg="white",
                  bd=0, padx=10, pady=2, cursor="hand2", command=self._reset).pack(side="right", padx=5)
        tk.Button(r2, text="🔍 Lancer la Recherche", font=FONT_HEAD, bg=C_PRIMARY, fg="white",
                  bd=0, padx=15, pady=2, cursor="hand2", command=self._search_reset_page).pack(side="right")

        # Main panel: sidebar + treeview
        main = tk.Frame(self, bg=C_BG)
        main.pack(fill="both", expand=True, padx=15)

        # Sidebar
        sb = tk.Frame(main, bg=C_SURFACE, padx=10, pady=10, width=185, bd=1, relief="solid")
        sb.pack(side="left", fill="y", pady=5)
        sb.pack_propagate(False)

        tk.Label(sb, text="ADMINISTRATION", font=FONT_HEAD, bg=C_SURFACE, fg=C_PRIMARY).pack(anchor="w", pady=5)
        self.btn_add  = self._sidebar_btn(sb, "➕ Nouvelle Notice",    C_PRIMARY_L, self._add_item)
        self.btn_edit = self._sidebar_btn(sb, "📝 Modifier Sélection", C_GOLD,      self._edit_item)
        self.btn_del  = self._sidebar_btn(sb, "🗑️ Supprimer Notice",   C_ACCENT,    self._delete_item)

        ttk.Separator(sb, orient="horizontal").pack(fill="x", pady=10)
        tk.Label(sb, text="ACTIONS EN MASSE", font=FONT_HEAD, bg=C_SURFACE, fg=C_PRIMARY).pack(anchor="w", pady=5)
        self.btn_import = self._sidebar_btn(sb, "📥 Import Données",   C_PRIMARY, self._import_bulk)
        self.btn_export = self._sidebar_btn(sb, "📤 Export Affichage", C_SUCCESS, self._export_view)

        # All admin buttons start disabled until logged in
        for btn in [self.btn_add, self.btn_edit, self.btn_del, self.btn_import, self.btn_export]:
            btn.config(state="disabled")

        # Treeview area
        tc = tk.Frame(main, bg=C_BG)
        tc.pack(side="right", fill="both", expand=True, padx=(10, 0))

        # Pagination bar
        pf = tk.Frame(tc, bg=C_BG)
        pf.pack(fill="x", pady=2)
        self.lbl_pag  = tk.Label(pf, text="Page 1 — 0 enregistrement(s)", font=FONT_BODY, bg=C_BG)
        self.lbl_pag.pack(side="left")
        self.btn_next = tk.Button(pf, text="Suivant ➡️",   font=FONT_SMALL, bg=C_BORDER, bd=0,
                                   padx=6, command=self._next_page, state="disabled")
        self.btn_next.pack(side="right", padx=2)
        self.btn_prev = tk.Button(pf, text="⬅️ Précédent", font=FONT_SMALL, bg=C_BORDER, bd=0,
                                   padx=6, command=self._prev_page, state="disabled")
        self.btn_prev.pack(side="right", padx=2)

        # Treeview with scrollbars
        scroll_y = tk.Scrollbar(tc, orient="vertical")
        scroll_x = tk.Scrollbar(tc, orient="horizontal")
        self.tree = ttk.Treeview(tc, columns=COLUMNS, show="headings",
                                  yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)
        scroll_y.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)
        scroll_x.pack(fill="x")

        for col in COLUMNS:
            self.tree.heading(col, text=COL_LABELS[col], anchor="w",
                              command=lambda c=col: self._sort_by_column(c))
            self.tree.column(col, width=COL_WIDTHS[col], minwidth=40, anchor="w")

        self.tree.bind("<Double-1>", lambda e: self._edit_item())

        # Status bar
        sb2 = tk.Frame(self, bg=C_BORDER, pady=3, padx=10)
        sb2.pack(fill="x", side="bottom")
        tk.Label(sb2, textvariable=self._status_var, font=FONT_SMALL, bg=C_BORDER, fg=C_TEXT).pack(side="left")
        tk.Label(sb2, textvariable=self._stats_var,  font=FONT_SMALL, bg=C_BORDER, fg=C_PRIMARY).pack(side="right")

    def _sidebar_btn(self, parent, text, color, command):
        btn = tk.Button(parent, text=text, font=FONT_BODY, bg=color, fg="white",
                        bd=0, pady=6, cursor="hand2", command=command)
        btn.pack(fill="x", pady=3)
        return btn

    # ── Login flow ────────────────────────────────────────────────────────────

    def _prompt_login(self):
        ConnectDialog(self, on_connect=self._on_login_success)

    def _on_login_success(self, dao_instance):
        self.dao = dao_instance
        self.lbl_status_badge.config(text="CONNECTÉ", bg=C_ONLINE)
        self._status_var.set(
            f"Connecté — {self.dao.host}:{self.dao.port}  |  Schéma : {self.dao.database}"
        )

        # Enable admin buttons
        for btn in [self.btn_add, self.btn_edit, self.btn_del, self.btn_import, self.btn_export]:
            btn.config(state="normal")

        # Create schema only if it doesn't exist yet (preserves existing data)
        try:
            if not self.dao.schema_exists():
                self.dao.init_schema()
                self._log("Schéma créé automatiquement (première initialisation).")
            else:
                self._log("Schéma existant détecté — chargement des données.")
        except Exception as e:
            self._log(f"Erreur schéma : {e}")

        # Load stats and data immediately
        self._refresh_stats()
        self._search()

    # ── Search & pagination ───────────────────────────────────────────────────

    def _reset(self):
        self.var_kw.set("")
        self.var_field.set("TOUS")
        self.var_fonds.set("TOUS")
        self.var_yr1.set("")
        self.var_yr2.set("")
        self._page = 0
        self._search()

    def _search_reset_page(self):
        self._page = 0
        self._search()

    def _search(self):
        if not self.dao:
            return
        self._status("Chargement des données...")

        kw     = self.var_kw.get().strip()
        fld    = self.var_field.get()
        fnds   = self.var_fonds.get()
        y1     = self.var_yr1.get().strip()
        y2     = self.var_yr2.get().strip()
        offset = self._page * self._limit

        def bg_worker():
            try:
                rows, total = self.dao.search(
                    keyword=kw, field=fld, fonds=fnds,
                    yr_from=y1, yr_to=y2,
                    sort_col=self._sort_col, sort_asc=self._sort_asc,
                    limit=self._limit, offset=offset,
                )
                self.after(0, self._update_table_view, rows, total, offset)
            except Exception as err:
                self.after(0, lambda e=err: messagebox.showerror(
                    "Erreur Requête", f"Impossible de récupérer les données :\n{e}"
                ))

        threading.Thread(target=bg_worker, daemon=True).start()

    def _update_table_view(self, rows, total, offset):
        self.tree.delete(*self.tree.get_children())

        for idx, r in enumerate(rows):
            tag  = "odd" if idx % 2 != 0 else "even"
            vals = ["" if r.get(c) is None else str(r.get(c)) for c in COLUMNS]
            iid  = self.tree.insert("", "end", values=vals, tags=(tag,))
            # Store hidden ID_NOTICE in the item's text field
            self.tree.item(iid, text=str(r.get("ID_NOTICE", "")))

        self.tree.tag_configure("odd",  background=C_ROW_ODD)
        self.tree.tag_configure("even", background=C_ROW_EVEN)

        end_idx   = min(offset + self._limit, total)
        start_idx = offset + 1 if total > 0 else 0
        self.lbl_pag.config(
            text=f"Page {self._page + 1} — {start_idx} à {end_idx} sur {total} enregistrement(s)"
        )
        self.btn_prev.config(state="normal" if self._page > 0 else "disabled")
        self.btn_next.config(state="normal" if end_idx < total else "disabled")
        self._status(f"Prêt — {total} enregistrement(s) trouvé(s).")

    def _sort_by_column(self, col):
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True
        self._search()

    def _next_page(self):
        self._page += 1
        self._search()

    def _prev_page(self):
        if self._page > 0:
            self._page -= 1
            self._search()

    def _refresh_stats(self):
        if not self.dao:
            return
        def bg():
            try:
                s = self.dao.get_stats()
                txt = (
                    f"Total notices : {s.get('total', 0)}  |  "
                    f"BUA (Arabe) : {s.get('bua', 0)}  |  "
                    f"BUF (Français) : {s.get('buf', 0)}  |  "
                    f"Auteurs : {s.get('auteurs', 0)}  |  "
                    f"Matières : {s.get('matieres', 0)}"
                )
                self.after(0, lambda: self._stats_var.set(txt))
            except Exception:
                pass
        threading.Thread(target=bg, daemon=True).start()

    # ── CRUD actions ──────────────────────────────────────────────────────────

    def _add_item(self):
        ItemDialog(self, self.dao, on_save=self._on_mutation_complete)

    def _edit_item(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sélection", "Veuillez sélectionner une ligne.")
            return
        id_notice = self.tree.item(sel[0], "text")
        try:
            record = self.dao.get_notice(id_notice)
            if record:
                ItemDialog(self, self.dao, record=record, on_save=self._on_mutation_complete)
            else:
                messagebox.showerror("Introuvable", "La notice n'existe plus en base.")
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def _delete_item(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sélection", "Veuillez sélectionner la ligne à supprimer.")
            return
        id_notice = self.tree.item(sel[0], "text")
        titre     = self.tree.item(sel[0], "values")[2]
        if messagebox.askyesno("Confirmation",
                               f"Supprimer définitivement la notice :\n\"{titre}\" ?"):
            try:
                self.dao.delete_notice(id_notice)
                self._log(f"Notice supprimée (ID: {id_notice}).")
                self._on_mutation_complete()
            except Exception as e:
                messagebox.showerror("Échec", f"Action impossible :\n{e}")

    def _on_mutation_complete(self):
        self._refresh_stats()
        self._search()

    # ── Bulk import (ETL pipeline) ────────────────────────────────────────────

    def _import_bulk(self):
        fp = filedialog.askopenfilename(
            filetypes=[("Fichiers supportés", "*.csv *.xls *.xlsx"), ("CSV", "*.csv")]
        )
        if not fp:
            return

        # Ask which fonds this file belongs to
        top = tk.Toplevel(self)
        top.title("Fonds cible")
        top.geometry("300x120")
        top.resizable(False, False)
        top.grab_set()
        tk.Label(top, text="Choisir le fonds de destination :", font=FONT_BODY, pady=10).pack()
        v_f = tk.StringVar(value="BUF")
        ttk.Combobox(top, textvariable=v_f, values=["BUA", "BUF"],
                     state="readonly", font=FONT_BODY).pack()

        def proceed():
            f_code = v_f.get()
            top.destroy()
            self._process_bulk_file(fp, f_code)

        tk.Button(top, text="Valider", font=FONT_HEAD, bg=C_PRIMARY, fg="white",
                  command=proceed, bd=0, padx=10, pady=2).pack(pady=10)

    def _process_bulk_file(self, filepath, fonds_code):
        if not PANDAS_OK:
            messagebox.showerror("Dépendance manquante",
                                 "pandas est requis pour l'import.\n"
                                 "pip install pandas openpyxl xlrd")
            return

        self._status("Lecture du fichier et préparation de l'import...")
        path = Path(filepath)

        # Progress window
        p_win = tk.Toplevel(self)
        p_win.title("Importation en cours...")
        p_win.geometry("420x110")
        p_win.resizable(False, False)
        p_win.grab_set()
        lbl_p = tk.Label(p_win, text="Initialisation...", font=FONT_BODY, pady=10)
        lbl_p.pack()
        p_bar = ttk.Progressbar(p_win, orient="horizontal", length=380, mode="determinate")
        p_bar.pack(pady=5)

        def run_import():
            try:
                # 1. EXTRACT
                self.after(0, lambda: lbl_p.config(text="Extraction des données..."))
                if path.suffix.lower() == ".csv":
                    df = export.export_data_csv(str(path))
                else:
                    df = export.export_data_excel(str(path))

                if df is None or df.empty:
                    self.after(0, lambda: messagebox.showerror("Vide", "Aucune donnée lue dans ce fichier."))
                    self.after(0, p_win.destroy)
                    return

                # 2. TRANSFORM
                self.after(0, lambda: lbl_p.config(text="Transformation ETL..."))
                lang_code    = self.dao.LANG_BY_FONDS.get(fonds_code, "fre")
                df_transform = transform.transform(df, lang_code)

                # 3. LOAD (prepare SQL-ready tables)
                self.after(0, lambda: lbl_p.config(text="Préparation des tables SQL..."))
                tables = load.export_for_sql(df_transform)

                # 4. INSERT into DB with progress updates
                def progress(msg, pct):
                    self.after(0, lambda: p_bar.config(value=pct))
                    self.after(0, lambda: lbl_p.config(text=msg))

                count = self.dao.load_etl_tables(tables, progress_cb=progress)

                self.after(0, p_win.destroy)
                self.after(0, lambda: messagebox.showinfo(
                    "Succès", f"Import terminé !\n{count} nouvelle(s) notice(s) insérée(s)."
                ))
                self.after(0, self._on_mutation_complete)

            except Exception as e:
                self.after(0, p_win.destroy)
                self.after(0, lambda: messagebox.showerror("Erreur d'import", f"Échec :\n{e}"))

        threading.Thread(target=run_import, daemon=True).start()

    # ── Export visible grid to CSV ────────────────────────────────────────────

    def _export_view(self):
        fp = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Fichier CSV", "*.csv")],
        )
        if not fp:
            return
        try:
            with open(fp, mode="w", encoding="utf-8-sig", newline="") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerow([COL_LABELS[c] for c in COLUMNS])
                for iid in self.tree.get_children():
                    writer.writerow(self.tree.item(iid, "values"))
            messagebox.showinfo("Exportation", "Données exportées avec succès.")
        except Exception as e:
            messagebox.showerror("Erreur Export", str(e))

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _status(self, msg):
        self._status_var.set(msg)

    def _log(self, msg):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    if not MYSQL_OK:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Dépendance critique",
            "Le pilote mysql-connector-python est introuvable.\n"
            "Installez-le : pip install mysql-connector-python",
        )
    else:
        app = MainWindow()
        app.mainloop()