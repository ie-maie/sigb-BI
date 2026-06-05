"""
╔══════════════════════════════════════════════════════════════════════════════╗
║        SIGB — Système Intégré de Gestion Bibliothécaire                     ║
║        Interface Python/Tkinter  ←→  MySQL                                 ║
║                                                                              ║
║  PRÉREQUIS (installer une seule fois) :                                      ║
║    pip install mysql-connector-python python-dotenv pandas xlrd openpyxl    ║
║                                                                              ║
║  CONFIGURATION :                                                             ║
║    1. Configurer DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASSWORD dans .env   ║
║    2. Verifier que la VM MySQL est accessible depuis ce poste                ║
║    3. Lancer l'interface depuis la racine du projet                          ║
║                                                                              ║
║  LANCEMENT :                                                                 ║
║    python UI.py                                                              ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import os
from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

ROOT_DIR = Path(__file__).resolve().parent
ENV_PATH = ROOT_DIR / ".env"
SCHEMA_PATH = ROOT_DIR / "sql" / "01_create_schema.sql"

if load_dotenv:
    load_dotenv(ENV_PATH)

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION CONFIG  ← Modifier ces valeurs ou le fichier .env
# ══════════════════════════════════════════════════════════════════════════════

DB_HOST        = os.getenv("DB_HOST", "localhost")
DB_PORT        = int(os.getenv("DB_PORT", "3306"))
DB_NAME        = os.getenv("DB_NAME", "sigb")
DB_USER        = os.getenv("DB_USER", "sigb_user")
DB_PASSWORD    = os.getenv("DB_PASSWORD", "")

# ══════════════════════════════════════════════════════════════════════════════
#  IMPORTS
# ══════════════════════════════════════════════════════════════════════════════

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import csv
import re
import unicodedata
from datetime import datetime

try:
    import mysql.connector
    MYSQL_OK = True
except ImportError:
    mysql = None
    MYSQL_OK = False

try:
    import pandas as pd
    PANDAS_OK = True
except ImportError:
    PANDAS_OK = False


# ══════════════════════════════════════════════════════════════════════════════
#  CONSTANTES & THÈME
# ══════════════════════════════════════════════════════════════════════════════

APP_TITLE   = "SIGB — MySQL"
APP_VERSION = "2.0"

C_BG        = "#F0F2F5"
C_SURFACE   = "#FFFFFF"
C_PRIMARY   = "#1A3A5C"
C_PRIMARY_L = "#2E6DA4"
C_ACCENT    = "#C0392B"
C_GOLD      = "#D4A017"
C_BUA       = "#1A3A5C"
C_BUF       = "#8B1A1A"
C_TEXT      = "#1A1A2E"
C_MUTED     = "#6B7A99"
C_BORDER    = "#DDE3EC"
C_ROW_ODD   = "#F7F9FC"
C_ROW_EVEN  = "#FFFFFF"
C_SEL       = "#D0E4F7"
C_SUCCESS   = "#2A6E1A"
C_ONLINE    = "#27AE60"
C_OFFLINE   = "#E74C3C"
C_ORANGE    = "#E67E22"

FONT_TITLE  = ("Segoe UI", 13, "bold")
FONT_HEAD   = ("Segoe UI", 10, "bold")
FONT_BODY   = ("Segoe UI", 9)
FONT_SMALL  = ("Segoe UI", 8)
FONT_MONO   = ("Courier New", 9)

# Colonnes affichées dans le tableau
COLUMNS = [
    "N_INVENTAIRE", "COTE", "TITRE",
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
    "N_INVENTAIRE": 75,
    "COTE":         105,
    "TITRE":        310,
    "NOM_AUTEUR":   170,
    "NOM_LIEU":     90,
    "NOM_EDITEUR":  120,
    "ANNEE":        58,
    "NB_PAGES":     60,
    "NOM_MATIERE":  200,
    "CODE_FONDS":   58,
}

# ══════════════════════════════════════════════════════════════════════════════
#  LEGACY ORACLE DDL (non utilisé) — conservé pour référence uniquement
#  NOTE: Le script DDL actif pour l'application est `sql/01_create_schema.sql` (MySQL).
# ══════════════════════════════════════════════════════════════════════════════

DDL_SCRIPT = """
-- ── Table FONDS ────────────────────────────────────────────────
CREATE TABLE FONDS (
    ID_FONDS   NUMBER(3)    GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    CODE       CHAR(3)      NOT NULL UNIQUE,
    LIBELLE    VARCHAR2(100) NOT NULL
);
INSERT INTO FONDS (CODE, LIBELLE) VALUES ('BUA', 'Bibliotheque universitaire arabe');
INSERT INTO FONDS (CODE, LIBELLE) VALUES ('BUF', 'Bibliotheque universitaire francaise');

-- ── Table AUTEUR ────────────────────────────────────────────────
CREATE TABLE AUTEUR (
    ID_AUTEUR  NUMBER(10)   GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    NOM        VARCHAR2(300) NOT NULL,
    CONSTRAINT UQ_AUTEUR_NOM UNIQUE (NOM)
);

-- ── Table EDITEUR ───────────────────────────────────────────────
CREATE TABLE EDITEUR (
    ID_EDITEUR NUMBER(10)   GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    NOM        VARCHAR2(200) NOT NULL,
    LIEU       VARCHAR2(100),
    CONSTRAINT UQ_EDITEUR UNIQUE (NOM, LIEU)
);

-- ── Table MATIERE ───────────────────────────────────────────────
CREATE TABLE MATIERE (
    ID_MATIERE NUMBER(10)   GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    LIBELLE    VARCHAR2(400) NOT NULL,
    CONSTRAINT UQ_MATIERE UNIQUE (LIBELLE)
);

-- ── Table NOTICE ────────────────────────────────────────────────
CREATE TABLE NOTICE (
    ID_NOTICE    NUMBER(10)   GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    N_INVENTAIRE NUMBER(10)   NOT NULL UNIQUE,
    COTE         VARCHAR2(120) NOT NULL,
    TITRE        VARCHAR2(600) NOT NULL,
    ANNEE        NUMBER(4),
    NB_PAGES     NUMBER(5),
    ID_FONDS     NUMBER(3)    NOT NULL,
    ID_AUTEUR    NUMBER(10),
    ID_EDITEUR   NUMBER(10),
    ID_MATIERE   NUMBER(10),
    DATE_AJOUT   TIMESTAMP    DEFAULT SYSTIMESTAMP,
    CONSTRAINT FK_NOT_FONDS   FOREIGN KEY (ID_FONDS)   REFERENCES FONDS(ID_FONDS),
    CONSTRAINT FK_NOT_AUTEUR  FOREIGN KEY (ID_AUTEUR)  REFERENCES AUTEUR(ID_AUTEUR),
    CONSTRAINT FK_NOT_EDITEUR FOREIGN KEY (ID_EDITEUR) REFERENCES EDITEUR(ID_EDITEUR),
    CONSTRAINT FK_NOT_MATIERE FOREIGN KEY (ID_MATIERE) REFERENCES MATIERE(ID_MATIERE)
);
CREATE INDEX IDX_NOTICE_TITRE ON NOTICE(TITRE);
CREATE INDEX IDX_NOTICE_ANNEE ON NOTICE(ANNEE);
CREATE INDEX IDX_NOTICE_COTE  ON NOTICE(COTE);

-- ── Table EXEMPLAIRE ────────────────────────────────────────────
CREATE TABLE EXEMPLAIRE (
    ID_EXEMPLAIRE NUMBER(10)  GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    ID_NOTICE     NUMBER(10)  NOT NULL,
    STATUT        VARCHAR2(20) DEFAULT 'disponible'
                  CHECK (STATUT IN ('disponible','emprunte','perdu','reserve')),
    LOCALISATION  VARCHAR2(100),
    CONSTRAINT FK_EX_NOTICE FOREIGN KEY (ID_NOTICE) REFERENCES NOTICE(ID_NOTICE)
);

-- ── Vue v_notices ────────────────────────────────────────────────
CREATE OR REPLACE VIEW LEGACY_NOTICES AS
SELECT
    n.ID_NOTICE,
    n.N_INVENTAIRE,
    n.COTE,
    n.TITRE,
    a.NOM        AS NOM_AUTEUR,
    e.LIEU       AS NOM_LIEU,
    e.NOM        AS NOM_EDITEUR,
    n.ANNEE,
    n.NB_PAGES,
    m.LIBELLE    AS NOM_MATIERE,
    f.CODE       AS CODE_FONDS,
    n.DATE_AJOUT
FROM NOTICE n
LEFT JOIN AUTEUR  a ON a.ID_AUTEUR  = n.ID_AUTEUR
LEFT JOIN EDITEUR e ON e.ID_EDITEUR = n.ID_EDITEUR
LEFT JOIN MATIERE m ON m.ID_MATIERE = n.ID_MATIERE
JOIN      FONDS   f ON f.ID_FONDS   = n.ID_FONDS;
"""


# ══════════════════════════════════════════════════════════════════════════════
#  COUCHE ACCÈS DONNÉES (DAO)
# ══════════════════════════════════════════════════════════════════════════════

class MySQLDAO:
    """
    Toutes les interactions avec la base MySQL passent par cette classe.
    Le schema correspond aux tables creees dans sql/01_create_schema.sql.
    """

    LANG_BY_FONDS = {"BUA": "ara", "BUF": "fre"}
    FONDS_BY_LANG = {"ara": "BUA", "fre": "BUF"}

    def __init__(self, host=DB_HOST, port=DB_PORT, database=DB_NAME,
                 user=DB_USER, password=DB_PASSWORD):
        self.host       = host
        self.port       = int(port)
        self.database   = database
        self.user       = user
        self.password   = password
        self.conn       = None
        self._connected = False

    # ── Connexion ──────────────────────────────────────────────────────────

    def connect(self):
        """Etablit la connexion MySQL."""
        if not MYSQL_OK:
            raise RuntimeError(
                "mysql-connector-python non installe.\n"
                "Executez : pip install mysql-connector-python python-dotenv"
            )
        self.conn = mysql.connector.connect(
            host=self.host,
            port=self.port,
            database=self.database,
            user=self.user,
            password=self.password,
            charset="utf8mb4",
            connection_timeout=10,
            use_pure=True,
        )
        if not self.conn.is_connected():
            raise RuntimeError("La connexion MySQL a ete creee mais elle n'est pas active.")
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
        """Vérifie que la connexion est toujours active."""
        try:
            self.conn.ping(reconnect=True, attempts=1, delay=0)
            return True
        except Exception:
            self._connected = False
            return False

    # ── Initialisation du schéma ──────────────────────────────────────────

    def init_schema(self):
        """Crée les tables si elles n'existent pas encore."""
        cur = self.conn.cursor()
        cur.execute("SHOW TABLES LIKE 'notice'")
        if cur.fetchone() is None:
            if not SCHEMA_PATH.exists():
                raise FileNotFoundError(f"Schema SQL introuvable : {SCHEMA_PATH}")
            for stmt in SCHEMA_PATH.read_text(encoding="utf-8").split(";"):
                stmt = stmt.strip()
                if stmt and not stmt.startswith("--"):
                    cur.execute(stmt)
            self.conn.commit()
            cur.close()
            return True
        cur.close()
        return False

    # ── RECHERCHE ─────────────────────────────────────────────────────────

    def search(self, keyword="", field="TOUS", fonds="TOUS",
               yr_from=None, yr_to=None,
               sort_col="N_INVENTAIRE", sort_asc=True,
               limit=500, offset=0):
        """
        Recherche paginee dans les tables MySQL normalisees.
        Retourne (rows: list[dict], total: int)
        """
        field_map = {
            "TOUS":       None,
            "TITRE":      "TITRE",
            "AUTEUR":     "NOM_AUTEUR",
            "COTE":       "COTE",
            "MATIERE":    "NOM_MATIERE",
            "EDITEUR":    "NOM_EDITEUR",
            "LIEU":       "NOM_LIEU",
        }
        sql_field = field_map.get(field.upper())

        conditions = []
        params     = []

        if keyword:
            kw = f"%{keyword.upper()}%"
            if sql_field:
                conditions.append(f"UPPER(COALESCE({sql_field}, '')) LIKE %s")
                params.append(kw)
            else:
                conditions.append(
                    "(UPPER(COALESCE(TITRE, '')) LIKE %s "
                    "OR UPPER(COALESCE(NOM_AUTEUR, '')) LIKE %s "
                    "OR UPPER(COALESCE(COTE, '')) LIKE %s "
                    "OR UPPER(COALESCE(NOM_MATIERE, '')) LIKE %s "
                    "OR UPPER(COALESCE(NOM_EDITEUR, '')) LIKE %s)"
                )
                params.extend([kw] * 5)

        if fonds and fonds != "TOUS":
            conditions.append("CODE_FONDS = %s")
            params.append(fonds)

        if yr_from:
            conditions.append("ANNEE >= %s")
            params.append(int(yr_from))
        if yr_to:
            conditions.append("ANNEE <= %s")
            params.append(int(yr_to))

        where = ("WHERE " + " AND ".join(conditions)) if conditions else ""
        direction = "ASC" if sort_asc else "DESC"
        safe_sort = sort_col if sort_col in COL_LABELS else "N_INVENTAIRE"

        base_sql = self._base_select()
        sql = f"""
            SELECT *
            FROM ({base_sql}) v
            {where}
            ORDER BY ({safe_sort} IS NULL), {safe_sort} {direction}
            LIMIT %s OFFSET %s
        """
        sql_count = f"SELECT COUNT(*) FROM ({base_sql}) v {where}"

        # Use a fresh connection for both the count and the paged select to avoid using
        # the main connection from a background thread (thread-safety issues with C ext).
        conn2 = mysql.connector.connect(
            host=self.host, port=self.port, database=self.database,
            user=self.user, password=self.password, charset="utf8mb4",
            connection_timeout=10, use_pure=True,
        )
        try:
            cur2 = conn2.cursor()
            cur2.execute(sql_count, tuple(params))
            total = cur2.fetchone()[0]
            cur2.close()

            curd = conn2.cursor(dictionary=True)
            curd.execute(sql, tuple(params + [int(limit), int(offset)]))
            rows = curd.fetchall()
            curd.close()
        finally:
            try:
                conn2.close()
            except Exception:
                pass
        return rows, total

    def _base_select(self):
        return """
            SELECT
                n.id_notice AS ID_NOTICE,
                ex.num_inventaire AS N_INVENTAIRE,
                COALESCE(ex.cote_exemplaire, c.cote) AS COTE,
                n.titre AS TITRE,
                GROUP_CONCAT(
                    DISTINCT TRIM(CONCAT(COALESCE(a.nom, ''), ' ', COALESCE(a.prenom, '')))
                    SEPARATOR '; '
                ) AS NOM_AUTEUR,
                ed.ville AS NOM_LIEU,
                ed.nom_editeur AS NOM_EDITEUR,
                n.annee_pub AS ANNEE,
                n.nb_pages AS NB_PAGES,
                GROUP_CONCAT(DISTINCT m.libelle SEPARATOR '; ') AS NOM_MATIERE,
                CASE l.code_langue
                    WHEN 'ara' THEN 'BUA'
                    WHEN 'fre' THEN 'BUF'
                    ELSE UPPER(l.code_langue)
                END AS CODE_FONDS,
                n.date_catalogage AS DATE_AJOUT
            FROM notice n
            LEFT JOIN langue l ON l.id_langue = n.id_langue
            LEFT JOIN classification c ON c.id_classification = n.id_classification
            LEFT JOIN editeur ed ON ed.id_editeur = n.id_editeur
            LEFT JOIN exemplaire ex ON ex.id_notice = n.id_notice
            LEFT JOIN notice_auteur na ON na.id_notice = n.id_notice
            LEFT JOIN auteur a ON a.id_auteur = na.id_auteur
            LEFT JOIN notice_matiere nm ON nm.id_notice = n.id_notice
            LEFT JOIN matiere m ON m.id_matiere = nm.id_matiere
            GROUP BY
                n.id_notice, ex.num_inventaire, ex.cote_exemplaire, c.cote,
                n.titre, ed.ville, ed.nom_editeur, n.annee_pub, n.nb_pages,
                l.code_langue, n.date_catalogage
        """

    # ── OBTENIR UNE NOTICE COMPLÈTE ──────────────────────────────────────

    def get_notice(self, id_notice):
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

    # ── INSERT ou GET FK (pour auteur/editeur/matiere) ───────────────────

    def _clean(self, value):
        if value is None:
            return ""
        text = str(value).strip()
        return "".join(c for c in text if unicodedata.category(c)[0] != "C").strip()

    def _safe_int(self, value):
        value = self._clean(value)
        try:
            return int(float(value)) if value else None
        except (TypeError, ValueError):
            return None

    def _split_values(self, value):
        value = self._clean(value)
        if not value:
            return []
        return [p.strip() for p in re.split(r"\s*[;/]\s*", value) if p.strip()]

    def _parse_auteur(self, value):
        value = self._clean(value)
        if not value:
            return None, None
        if "," in value:
            nom, prenom = value.split(",", 1)
            return self._clean(nom), self._clean(prenom) or None
        parts = value.split()
        if len(parts) > 1 and parts[0].isupper():
            return parts[0], " ".join(parts[1:]) or None
        return value, None

    def _get_langue_id(self, fonds_code):
        code = self.LANG_BY_FONDS.get((fonds_code or "BUF").upper(), "fre")
        cur = self.conn.cursor()
        cur.execute("SELECT id_langue FROM langue WHERE code_langue=%s", (code,))
        row = cur.fetchone()
        if row:
            cur.close()
            return row[0]
        libelle = {"ara": "Arabe", "fre": "Francais", "eng": "Anglais"}.get(code, code)
        cur.execute(
            "INSERT INTO langue (code_langue, libelle) VALUES (%s, %s)",
            (code, libelle),
        )
        self.conn.commit()
        new_id = cur.lastrowid
        cur.close()
        return new_id

    def _get_or_create_classification(self, cote):
        cote = self._clean(cote)
        if not cote:
            return None
        cur = self.conn.cursor()
        cur.execute("INSERT IGNORE INTO classification (cote) VALUES (%s)", (cote,))
        cur.execute("SELECT id_classification FROM classification WHERE cote=%s", (cote,))
        row = cur.fetchone()
        cur.close()
        return row[0] if row else None

    def _get_or_create_matiere(self, libelle):
        libelle = self._clean(libelle)
        if not libelle:
            return None
        cur = self.conn.cursor()
        cur.execute("INSERT IGNORE INTO matiere (libelle) VALUES (%s)", (libelle,))
        cur.execute("SELECT id_matiere FROM matiere WHERE libelle=%s", (libelle,))
        row = cur.fetchone()
        cur.close()
        return row[0] if row else None

    def _get_or_create_auteur(self, auteur):
        nom, prenom = self._parse_auteur(auteur)
        if not nom:
            return None
        cur = self.conn.cursor()
        cur.execute(
            "INSERT IGNORE INTO auteur (nom, prenom) VALUES (%s, %s)",
            (nom, prenom),
        )
        cur.execute(
            "SELECT id_auteur FROM auteur "
            "WHERE nom=%s AND COALESCE(prenom,'')=COALESCE(%s,'')",
            (nom, prenom),
        )
        row = cur.fetchone()
        cur.close()
        return row[0] if row else None

    def _get_or_create_editeur(self, nom, lieu):
        nom = self._clean(nom)
        lieu = self._clean(lieu) or None
        if not nom:
            return None
        cur = self.conn.cursor()
        cur.execute(
            "INSERT IGNORE INTO editeur (nom_editeur, ville) VALUES (%s, %s)",
            (nom, lieu),
        )
        cur.execute(
            "SELECT id_editeur FROM editeur "
            "WHERE nom_editeur=%s AND COALESCE(ville,'')=COALESCE(%s,'')",
            (nom, lieu),
        )
        row = cur.fetchone()
        cur.close()
        return row[0] if row else None

    def _link_people_and_subjects(self, cur, id_notice, auteurs, matieres):
        for auteur in self._split_values(auteurs):
            id_auteur = self._get_or_create_auteur(auteur)
            if id_auteur:
                cur.execute(
                    "INSERT IGNORE INTO notice_auteur (id_notice, id_auteur) VALUES (%s, %s)",
                    (id_notice, id_auteur),
                )
        for matiere in self._split_values(matieres):
            id_matiere = self._get_or_create_matiere(matiere)
            if id_matiere:
                cur.execute(
                    "INSERT IGNORE INTO notice_matiere (id_notice, id_matiere) VALUES (%s, %s)",
                    (id_notice, id_matiere),
                )

    def _upsert_exemplaire(self, cur, id_notice, inventaire, cote):
        inventaire = self._clean(inventaire)
        cote = self._clean(cote)
        if not inventaire:
            return
        cur.execute("SELECT id_exemplaire FROM exemplaire WHERE id_notice=%s LIMIT 1", (id_notice,))
        row = cur.fetchone()
        if row:
            cur.execute(
                "UPDATE exemplaire SET num_inventaire=%s, cote_exemplaire=%s WHERE id_exemplaire=%s",
                (inventaire, cote, row[0]),
            )
        else:
            cur.execute(
                "INSERT INTO exemplaire (num_inventaire, id_notice, cote_exemplaire) VALUES (%s, %s, %s)",
                (inventaire, id_notice, cote),
            )

    def _notice_exists_by_inventory(self, inventaire):
        inventaire = self._clean(inventaire)
        if not inventaire:
            return False
        cur = self.conn.cursor()
        cur.execute("SELECT 1 FROM exemplaire WHERE num_inventaire=%s LIMIT 1", (inventaire,))
        row = cur.fetchone()
        cur.close()
        return bool(row)

    # ── AJOUTER UNE NOTICE ───────────────────────────────────────────────

    def add_notice(self, data):
        """
        data = dict avec clés :
          N_INVENTAIRE, COTE, TITRE, NOM_AUTEUR, NOM_LIEU,
          NOM_EDITEUR, ANNEE, NB_PAGES, NOM_MATIERE, CODE_FONDS
        """
        inv = self._clean(data.get("N_INVENTAIRE"))
        if self._notice_exists_by_inventory(inv):
            raise RuntimeError(f"Le numero d'inventaire existe deja : {inv}")

        cur = self.conn.cursor()
        cur.execute("""
            INSERT INTO notice
                (titre, annee_pub, nb_pages, id_editeur, id_langue, id_classification)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (
            self._clean(data.get("TITRE")),
            self._safe_int(data.get("ANNEE")),
            self._safe_int(data.get("NB_PAGES")),
            self._get_or_create_editeur(data.get("NOM_EDITEUR"), data.get("NOM_LIEU")),
            self._get_langue_id(data.get("CODE_FONDS", "BUF")),
            self._get_or_create_classification(data.get("COTE")),
        ))
        id_notice = cur.lastrowid
        self._link_people_and_subjects(
            cur, id_notice, data.get("NOM_AUTEUR"), data.get("NOM_MATIERE")
        )
        self._upsert_exemplaire(cur, id_notice, inv, data.get("COTE"))
        self.conn.commit()
        cur.close()

    # ── MODIFIER UNE NOTICE ──────────────────────────────────────────────

    def update_notice(self, id_notice, data):
        cur = self.conn.cursor()
        cur.execute("""
            UPDATE notice SET
                titre=%s,
                annee_pub=%s,
                nb_pages=%s,
                id_editeur=%s,
                id_langue=%s,
                id_classification=%s
            WHERE id_notice=%s
        """, (
            self._clean(data.get("TITRE")),
            self._safe_int(data.get("ANNEE")),
            self._safe_int(data.get("NB_PAGES")),
            self._get_or_create_editeur(data.get("NOM_EDITEUR"), data.get("NOM_LIEU")),
            self._get_langue_id(data.get("CODE_FONDS", "BUF")),
            self._get_or_create_classification(data.get("COTE")),
            id_notice,
        ))
        cur.execute("DELETE FROM notice_auteur WHERE id_notice=%s", (id_notice,))
        cur.execute("DELETE FROM notice_matiere WHERE id_notice=%s", (id_notice,))
        self._link_people_and_subjects(
            cur, id_notice, data.get("NOM_AUTEUR"), data.get("NOM_MATIERE")
        )
        self._upsert_exemplaire(
            cur, id_notice, data.get("N_INVENTAIRE"), data.get("COTE")
        )
        self.conn.commit()
        cur.close()

    # ── SUPPRIMER UNE NOTICE ─────────────────────────────────────────────

    def delete_notice(self, id_notice):
        cur = self.conn.cursor()
        cur.execute("DELETE FROM exemplaire WHERE id_notice=%s", (id_notice,))
        cur.execute("DELETE FROM notice WHERE id_notice=%s", (id_notice,))
        self.conn.commit()
        cur.close()

    # ── STATISTIQUES ─────────────────────────────────────────────────────

    def get_stats(self):
        # Use a dedicated connection for stats to avoid threading issues
        conn2 = mysql.connector.connect(
            host=self.host, port=self.port, database=self.database,
            user=self.user, password=self.password, charset="utf8mb4",
            connection_timeout=10, use_pure=True,
        )
        stats = {}
        try:
            cur = conn2.cursor()
            cur.execute("SELECT COUNT(*) FROM notice")
            stats["total"] = cur.fetchone()[0]
            cur.execute("""
                SELECT COUNT(*)
                FROM notice n
                JOIN langue l ON l.id_langue=n.id_langue
                WHERE l.code_langue='ara'
            """)
            stats["bua"] = cur.fetchone()[0]
            cur.execute("""
                SELECT COUNT(*)
                FROM notice n
                JOIN langue l ON l.id_langue=n.id_langue
                WHERE l.code_langue='fre'
            """)
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

    # ── IMPORT BULK (XLS / CSV) ──────────────────────────────────────────

    def bulk_insert(self, rows, fonds_code, progress_cb=None):
        """
        Insere une liste de dicts (colonnes normalisees).
        rows : liste de dicts avec clés Cote/Titre/Auteur/Lieu/Edition/Annee/Nb_pages/Matiere/Inventaire
        """
        id_langue = self._get_langue_id(fonds_code)
        cur = self.conn.cursor()
        ok = 0
        for i, r in enumerate(rows):
            try:
                inv = self._clean(r.get("Inventaire"))
                if not inv:
                    continue
                cur.execute("SELECT 1 FROM exemplaire WHERE num_inventaire=%s LIMIT 1", (inv,))
                if cur.fetchone():
                    continue
                titre = self._clean(r.get("Titre"))
                if not titre:
                    continue
                cote = self._clean(r.get("Cote"))
                cur.execute("""
                    INSERT INTO notice
                        (titre, annee_pub, nb_pages, id_editeur, id_langue, id_classification)
                    VALUES (%s, %s, %s, %s, %s, %s)
                """, (
                    titre,
                    self._safe_int(r.get("Annee")),
                    self._safe_int(r.get("Nb_pages")),
                    self._get_or_create_editeur(r.get("Edition"), r.get("Lieu")),
                    id_langue,
                    self._get_or_create_classification(cote),
                ))
                id_notice = cur.lastrowid
                self._link_people_and_subjects(
                    cur, id_notice, r.get("Auteur"), r.get("Matiere")
                )
                self._upsert_exemplaire(cur, id_notice, inv, cote)
                ok += 1
            except Exception:
                pass
            if progress_cb and i % 100 == 0:
                progress_cb(i, len(rows))
        self.conn.commit()
        cur.close()
        return ok


# ══════════════════════════════════════════════════════════════════════════════
#  FENÊTRE DE CONNEXION
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
        # Header
        hdr = tk.Frame(self, bg=C_PRIMARY, pady=14, padx=20)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🔐  Connexion MySQL",
                 font=FONT_TITLE, fg="white", bg=C_PRIMARY).pack(anchor="w")
        tk.Label(hdr, text="Renseignez vos identifiants ou modifiez le fichier .env",
                 font=FONT_SMALL, fg=C_GOLD, bg=C_PRIMARY).pack(anchor="w", pady=(4,0))

        body = tk.Frame(self, bg=C_BG, padx=24, pady=20)
        body.pack(fill="both")

        fields = [
            ("Hote MySQL :",          "host",     DB_HOST,     False),
            ("Port :",                "port",     str(DB_PORT), False),
            ("Base :",                "database", DB_NAME,     False),
            ("Utilisateur :",         "user",     DB_USER,     False),
            ("Mot de passe :",        "password", DB_PASSWORD, True),
        ]
        self.vars = {}
        for label, key, default, is_pass in fields:
            frm = tk.Frame(body, bg=C_BG)
            frm.pack(fill="x", pady=4)
            tk.Label(frm, text=label, font=FONT_SMALL, fg=C_MUTED,
                     bg=C_BG, width=22, anchor="w").pack(side="left")
            var = tk.StringVar(value=default)
            self.vars[key] = var
            e = ttk.Entry(frm, textvariable=var, width=40, font=FONT_BODY,
                          show="●" if is_pass else "")
            e.pack(side="left", fill="x", expand=True)
        tk.Label(body,
                 text="💡 Les valeurs par defaut viennent du fichier .env a la racine du projet.",
                 font=FONT_SMALL, fg=C_MUTED, bg=C_BG).pack(anchor="w", pady=(4,0))

        # Boutons
        btn_frm = tk.Frame(self, bg=C_SURFACE, pady=12, padx=24)
        btn_frm.pack(fill="x")
        ttk.Separator(btn_frm).pack(fill="x", pady=(0, 10))
        tk.Button(btn_frm, text="✕  Annuler", command=self.destroy,
                  font=FONT_BODY, bg=C_BG, fg=C_TEXT, relief="flat",
                  padx=14, pady=6, cursor="hand2").pack(side="right", padx=(8,0))
        tk.Button(btn_frm, text="🔌  Se connecter", command=self._connect,
                  font=FONT_BODY, bg=C_PRIMARY, fg="white",
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(side="right")

    def _connect(self):
        self.on_connect(
            host=self.vars["host"].get().strip(),
            port=self.vars["port"].get().strip(),
            database=self.vars["database"].get().strip(),
            user=self.vars["user"].get().strip(),
            password=self.vars["password"].get(),
        )
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
#  DIALOGUE AJOUT / MODIFICATION
# ══════════════════════════════════════════════════════════════════════════════

class NoticeDialog(tk.Toplevel):

    FIELDS = [
        ("N° Inventaire *",  "N_INVENTAIRE"),
        ("Cote *",            "COTE"),
        ("Titre *",           "TITRE"),
        ("Auteur",            "NOM_AUTEUR"),
        ("Lieu d'édition",    "NOM_LIEU"),
        ("Éditeur",           "NOM_EDITEUR"),
        ("Année",             "ANNEE"),
        ("Nb pages",          "NB_PAGES"),
        ("Matière",           "NOM_MATIERE"),
    ]

    def __init__(self, parent, title="Notice", data=None, on_save=None):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.configure(bg=C_BG)
        self.on_save = on_save
        self.entries = {}
        self._build(data or {})
        self.grab_set()
        self.focus_force()
        self.wait_window()

    def _build(self, data):
        # Header
        hdr = tk.Frame(self, bg=C_PRIMARY, pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text=self.title(), font=FONT_TITLE,
                 fg="white", bg=C_PRIMARY).pack(anchor="w")

        body = tk.Frame(self, bg=C_BG, padx=20, pady=16)
        body.pack(fill="both", expand=True)

        for i, (label, key) in enumerate(self.FIELDS):
            row, col = divmod(i, 2)
            colspan = 2 if key in ("TITRE", "NOM_MATIERE") else 1
            frm = tk.Frame(body, bg=C_BG)
            if colspan == 2:
                frm.grid(row=row, column=0, columnspan=4,
                         padx=0, pady=5, sticky="ew")
                w = 62
            else:
                frm.grid(row=row, column=col * 2, columnspan=2,
                         padx=(0, 16 if col == 0 else 0),
                         pady=5, sticky="ew")
                w = 30
            tk.Label(frm, text=label, font=FONT_SMALL,
                     fg=C_MUTED, bg=C_BG).pack(anchor="w")
            e = ttk.Entry(frm, width=w, font=FONT_BODY)
            val = data.get(key, "")
            e.insert(0, str(val) if val is not None else "")
            e.pack(fill="x")
            self.entries[key] = e

        # Fonds
        r_fonds = len(self.FIELDS) // 2 + 1
        frm_f = tk.Frame(body, bg=C_BG)
        frm_f.grid(row=r_fonds, column=0, columnspan=2, pady=5, sticky="ew")
        tk.Label(frm_f, text="Fonds *", font=FONT_SMALL,
                 fg=C_MUTED, bg=C_BG).pack(anchor="w")
        self.var_fonds = tk.StringVar(value=data.get("CODE_FONDS", "BUF"))
        ttk.Combobox(frm_f, textvariable=self.var_fonds,
                     values=["BUF", "BUA"], state="readonly",
                     width=10, font=FONT_BODY).pack(anchor="w")

        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.columnconfigure(2, weight=1)
        body.columnconfigure(3, weight=1)

        # Boutons
        btn_frm = tk.Frame(self, bg=C_SURFACE, pady=12, padx=20)
        btn_frm.pack(fill="x")
        ttk.Separator(btn_frm).pack(fill="x", pady=(0, 10))
        tk.Button(btn_frm, text="✕  Annuler", command=self.destroy,
                  font=FONT_BODY, bg=C_BG, fg=C_TEXT,
                  relief="flat", padx=14, pady=6,
                  cursor="hand2").pack(side="right", padx=(8, 0))
        tk.Button(btn_frm, text="✔  Enregistrer dans MySQL",
                  command=self._save,
                  font=FONT_BODY, bg=C_PRIMARY, fg="white",
                  relief="flat", padx=14, pady=6,
                  cursor="hand2").pack(side="right")

    def _save(self):
        inv   = self.entries["N_INVENTAIRE"].get().strip()
        cote  = self.entries["COTE"].get().strip()
        titre = self.entries["TITRE"].get().strip()
        if not inv or not cote or not titre:
            messagebox.showwarning("Champs obligatoires",
                                   "N° Inventaire, Cote et Titre sont obligatoires.",
                                   parent=self)
            return
        result = {key: self.entries[key].get().strip() for key in self.entries}
        result["CODE_FONDS"] = self.var_fonds.get()
        if self.on_save:
            self.on_save(result)
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
#  APPLICATION PRINCIPALE
# ══════════════════════════════════════════════════════════════════════════════

class SIGBApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.state("zoomed")
        self.configure(bg=C_BG)
        self.minsize(1100, 650)

        self.dao          = None
        self._rows        = []       # liste de dicts (résultats courants)
        self._total       = 0        # total MySQL
        self._page        = 0        # offset pagination
        self._page_size   = 100
        self._sort_col    = "N_INVENTAIRE"
        self._sort_asc    = True
        self._sel_id      = None     # ID_NOTICE sélectionné
        self._sel_data    = {}

        self._build_styles()
        self._build_menu()
        self._build_toolbar()
        self._build_main()
        self._build_statusbar()

        # Tentative de connexion automatique au démarrage
        self.after(400, self._auto_connect)

    # ── Styles ────────────────────────────────────────────────────────────

    def _build_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame",       background=C_BG)
        s.configure("TLabel",       background=C_BG, foreground=C_TEXT, font=FONT_BODY)
        s.configure("TLabelframe",  background=C_BG, foreground=C_PRIMARY, font=FONT_HEAD)
        s.configure("TLabelframe.Label", background=C_BG, foreground=C_PRIMARY, font=FONT_HEAD)
        s.configure("TButton",      background=C_SURFACE, foreground=C_TEXT,
                    font=FONT_BODY, padding=(10,5), relief="flat", borderwidth=1)
        s.map("TButton", background=[("active", C_ROW_ODD)])
        s.configure("TCombobox",    font=FONT_BODY)
        s.configure("TEntry",       font=FONT_BODY, padding=4)
        s.configure("Treeview",
                    background=C_SURFACE, foreground=C_TEXT, rowheight=24,
                    font=FONT_BODY, fieldbackground=C_SURFACE, borderwidth=0)
        s.configure("Treeview.Heading",
                    background=C_PRIMARY, foreground="white",
                    font=FONT_HEAD, relief="flat", padding=(4,6))
        s.map("Treeview.Heading", background=[("active", C_PRIMARY_L)])
        s.map("Treeview",
              background=[("selected", C_SEL)],
              foreground=[("selected", C_TEXT)])
        s.configure("TNotebook",     background=C_BG, tabmargins=[2,2,0,0])
        s.configure("TNotebook.Tab", background=C_BORDER, foreground=C_TEXT,
                    font=FONT_BODY, padding=[12,5])
        s.map("TNotebook.Tab",
              background=[("selected", C_PRIMARY)],
              foreground=[("selected", "white")])

    # ── Menu ──────────────────────────────────────────────────────────────

    def _build_menu(self):
        mb = tk.Menu(self, bg=C_SURFACE, fg=C_TEXT, font=FONT_BODY)
        self.config(menu=mb)

        m_db = tk.Menu(mb, tearoff=0, bg=C_SURFACE, fg=C_TEXT, font=FONT_BODY)
        m_db.add_command(label="🔌  Connecter…",           command=self._open_connect_dialog)
        m_db.add_command(label="🔌  Déconnecter",          command=self._disconnect)
        m_db.add_separator()
        m_db.add_command(label="🗄  Initialiser le schéma", command=self._init_schema)
        m_db.add_separator()
        m_db.add_command(label="✕  Quitter",               command=self.quit)
        mb.add_cascade(label="Base de données", menu=m_db)

        m_notices = tk.Menu(mb, tearoff=0, bg=C_SURFACE, fg=C_TEXT, font=FONT_BODY)
        m_notices.add_command(label="➕  Ajouter",             command=self._add_notice)
        m_notices.add_command(label="✏️  Modifier",            command=self._edit_notice)
        m_notices.add_command(label="🗑  Supprimer",           command=self._delete_notice)
        m_notices.add_separator()
        m_notices.add_command(label="🔄  Actualiser",          command=self._search)
        mb.add_cascade(label="Notices", menu=m_notices)

        m_imp = tk.Menu(mb, tearoff=0, bg=C_SURFACE, fg=C_TEXT, font=FONT_BODY)
        m_imp.add_command(label="📂  Importer BUA (XLS)…",    command=self._import_bua)
        m_imp.add_command(label="📂  Importer BUF (CSV)…",    command=self._import_buf)
        m_imp.add_separator()
        m_imp.add_command(label="💾  Exporter CSV…",           command=self._export_csv)
        mb.add_cascade(label="Import / Export", menu=m_imp)

        m_help = tk.Menu(mb, tearoff=0, bg=C_SURFACE, fg=C_TEXT, font=FONT_BODY)
        m_help.add_command(label="ℹ️  À propos", command=self._about)
        mb.add_cascade(label="Aide", menu=m_help)

    # ── Barre d'outils ────────────────────────────────────────────────────

    def _build_toolbar(self):
        tb = tk.Frame(self, bg=C_PRIMARY, pady=6, padx=10)
        tb.pack(fill="x", side="top")

        tk.Label(tb, text="📚  SIGB  MySQL",
                 font=("Segoe UI", 11, "bold"),
                 fg="white", bg=C_PRIMARY).pack(side="left", padx=(0, 24))

        btns = [
            ("➕ Ajouter",     self._add_notice,    C_GOLD,    C_TEXT),
            ("✏️ Modifier",    self._edit_notice,   C_SURFACE, C_PRIMARY),
            ("🗑 Supprimer",   self._delete_notice, C_ACCENT,  "white"),
            ("🔄 Actualiser",  self._search,        C_SURFACE, C_PRIMARY),
            ("📂 Importer",    self._import_menu,   C_SURFACE, C_PRIMARY),
            ("💾 Export CSV",  self._export_csv,    C_SURFACE, C_PRIMARY),
        ]
        for txt, cmd, bg, fg in btns:
            tk.Button(tb, text=txt, command=cmd,
                      font=FONT_BODY, bg=bg, fg=fg,
                      relief="flat", padx=12, pady=4,
                      cursor="hand2").pack(side="left", padx=3)

        # Indicateur connexion
        self.conn_dot  = tk.Label(tb, text="●", font=("Segoe UI", 14),
                                   fg=C_OFFLINE, bg=C_PRIMARY)
        self.conn_dot.pack(side="right", padx=4)
        self.conn_label = tk.Label(tb, text="Déconnecté",
                                    font=FONT_SMALL, fg=C_GOLD, bg=C_PRIMARY)
        self.conn_label.pack(side="right", padx=(0, 2))

        self.lbl_count = tk.Label(tb, text="", font=FONT_SMALL,
                                   fg=C_GOLD, bg=C_PRIMARY)
        self.lbl_count.pack(side="right", padx=(0, 16))

    # ── Corps principal ───────────────────────────────────────────────────

    def _build_main(self):
        paned = tk.PanedWindow(self, orient="horizontal",
                               bg=C_BG, sashwidth=5, bd=0)
        paned.pack(fill="both", expand=True, padx=8, pady=8)

        left = tk.Frame(paned, bg=C_SURFACE, bd=1, relief="solid", width=270)
        left.pack_propagate(False)
        paned.add(left, minsize=230)
        self._build_left(left)

        right = tk.Frame(paned, bg=C_BG)
        paned.add(right, minsize=650)
        self._build_right(right)

    # ── Panneau gauche ────────────────────────────────────────────────────

    def _build_left(self, parent):
        hdr = tk.Frame(parent, bg=C_PRIMARY, pady=10, padx=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🔍  Recherche & Filtres",
                 font=FONT_HEAD, fg="white", bg=C_PRIMARY).pack(anchor="w")

        scroll = tk.Frame(parent, bg=C_SURFACE)
        scroll.pack(fill="both", expand=True, padx=12, pady=10)

        # ── Recherche ──
        frm_s = ttk.LabelFrame(scroll, text="  Recherche  ", padding=8)
        frm_s.pack(fill="x", pady=(0,10))

        tk.Label(frm_s, text="Mot-clé :", font=FONT_SMALL, fg=C_MUTED, bg=C_BG).pack(anchor="w")
        self.var_kw = tk.StringVar()
        e = ttk.Entry(frm_s, textvariable=self.var_kw, width=26, font=FONT_BODY)
        e.pack(fill="x", pady=(2,6))
        e.bind("<Return>", lambda _: self._search())

        tk.Label(frm_s, text="Dans le champ :", font=FONT_SMALL, fg=C_MUTED, bg=C_BG).pack(anchor="w")
        self.var_field = tk.StringVar(value="TOUS")
        ttk.Combobox(frm_s, textvariable=self.var_field,
                     values=["TOUS","TITRE","AUTEUR","COTE","MATIERE","EDITEUR","LIEU"],
                     state="readonly", width=24, font=FONT_BODY).pack(fill="x", pady=(2,0))

        tk.Button(frm_s, text="🔍  Rechercher", command=self._search,
                  font=FONT_BODY, bg=C_PRIMARY, fg="white",
                  relief="flat", pady=5, cursor="hand2").pack(fill="x", pady=(10,0))

        # ── Filtres ──
        frm_f = ttk.LabelFrame(scroll, text="  Filtres  ", padding=8)
        frm_f.pack(fill="x", pady=(0,10))

        tk.Label(frm_f, text="Fonds :", font=FONT_SMALL, fg=C_MUTED, bg=C_BG).pack(anchor="w")
        self.var_fonds = tk.StringVar(value="TOUS")
        ttk.Combobox(frm_f, textvariable=self.var_fonds,
                     values=["TOUS","BUA","BUF"],
                     state="readonly", width=24, font=FONT_BODY).pack(fill="x", pady=(2,8))

        tk.Label(frm_f, text="Année de :", font=FONT_SMALL, fg=C_MUTED, bg=C_BG).pack(anchor="w")
        yr_row = tk.Frame(frm_f, bg=C_BG)
        yr_row.pack(fill="x", pady=2)
        self.var_yr1 = tk.StringVar()
        self.var_yr2 = tk.StringVar()
        ttk.Entry(yr_row, textvariable=self.var_yr1, width=8, font=FONT_BODY).pack(side="left")
        tk.Label(yr_row, text=" à ", font=FONT_SMALL, fg=C_MUTED, bg=C_BG).pack(side="left")
        ttk.Entry(yr_row, textvariable=self.var_yr2, width=8, font=FONT_BODY).pack(side="left")

        tk.Button(scroll, text="↺  Réinitialiser", command=self._reset,
                  font=FONT_SMALL, bg=C_BORDER, fg=C_TEXT,
                  relief="flat", pady=5, cursor="hand2").pack(fill="x", pady=(4,12))

        # ── Stats ──
        frm_st = ttk.LabelFrame(scroll, text="  Statistiques MySQL  ", padding=8)
        frm_st.pack(fill="x")

        self.stats_vars = {}
        for label, key, color in [
            ("Total notices",  "total",    C_PRIMARY),
            ("Fonds BUA",      "bua",      C_BUA),
            ("Fonds BUF",      "buf",      C_BUF),
            ("Matières",       "matieres", C_GOLD),
            ("Auteurs",        "auteurs",  C_SUCCESS),
            ("Affichées",      "shown",    C_ORANGE),
        ]:
            row = tk.Frame(frm_st, bg=C_BG)
            row.pack(fill="x", pady=2)
            tk.Label(row, text=label, font=FONT_SMALL,
                     fg=C_MUTED, bg=C_BG).pack(side="left")
            var = tk.StringVar(value="—")
            self.stats_vars[key] = var
            tk.Label(row, textvariable=var, font=("Segoe UI", 9, "bold"),
                     fg=color, bg=C_BG).pack(side="right")

        tk.Button(scroll, text="🔄  Rafraîchir les stats", command=self._refresh_stats,
                  font=FONT_SMALL, bg=C_BORDER, fg=C_TEXT,
                  relief="flat", pady=5, cursor="hand2").pack(fill="x", pady=(8,0))

    # ── Panneau droit ─────────────────────────────────────────────────────

    def _build_right(self, parent):
        self.notebook = ttk.Notebook(parent)
        self.notebook.pack(fill="both", expand=True)

        tab1 = tk.Frame(self.notebook, bg=C_BG)
        self.notebook.add(tab1, text="  📋  Catalogue  ")
        self._build_table_tab(tab1)

        tab2 = tk.Frame(self.notebook, bg=C_BG)
        self.notebook.add(tab2, text="  📄  Détail notice  ")
        self._build_detail_tab(tab2)

        tab3 = tk.Frame(self.notebook, bg=C_BG)
        self.notebook.add(tab3, text="  💾  Import / Export  ")
        self._build_io_tab(tab3)

    # ── Onglet tableau ────────────────────────────────────────────────────

    def _build_table_tab(self, parent):
        frm = tk.Frame(parent, bg=C_BG)
        frm.pack(fill="both", expand=True, padx=4, pady=4)

        vsb = ttk.Scrollbar(frm, orient="vertical")
        hsb = ttk.Scrollbar(frm, orient="horizontal")

        self.tree = ttk.Treeview(
            frm, columns=COLUMNS, show="headings",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
            selectmode="browse",
        )
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        for col in COLUMNS:
            self.tree.heading(col, text=COL_LABELS[col],
                              command=lambda c=col: self._sort_by(c))
            self.tree.column(col, width=COL_WIDTHS[col],
                             minwidth=40, stretch=(col == "TITRE"))

        self.tree.tag_configure("odd",  background=C_ROW_ODD)
        self.tree.tag_configure("even", background=C_ROW_EVEN)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>",         self._on_double_click)
        self.tree.bind("<Delete>",           lambda _: self._delete_notice())

        # Menu contextuel
        ctx = tk.Menu(self, tearoff=0, bg=C_SURFACE, fg=C_TEXT, font=FONT_BODY)
        ctx.add_command(label="👁  Voir détail",  command=lambda: self.notebook.select(1))
        ctx.add_command(label="✏️  Modifier",     command=self._edit_notice)
        ctx.add_separator()
        ctx.add_command(label="🗑  Supprimer",    command=self._delete_notice)
        self.tree.bind("<Button-3>",
                       lambda e: (self.tree.selection_set(
                           self.tree.identify_row(e.y)),
                           ctx.post(e.x_root, e.y_root)))

        # Pagination
        pag_frm = tk.Frame(parent, bg=C_SURFACE, pady=6)
        pag_frm.pack(fill="x", padx=4)
        self.btn_prev = tk.Button(pag_frm, text="◀ Précédent",
                                   command=self._prev_page,
                                   font=FONT_SMALL, bg=C_BORDER,
                                   fg=C_TEXT, relief="flat",
                                   padx=10, pady=4, cursor="hand2")
        self.btn_prev.pack(side="left", padx=(4,2))
        self.lbl_page = tk.Label(pag_frm, text="Page —",
                                  font=FONT_SMALL, fg=C_MUTED, bg=C_SURFACE)
        self.lbl_page.pack(side="left", padx=8)
        self.btn_next = tk.Button(pag_frm, text="Suivant ▶",
                                   command=self._next_page,
                                   font=FONT_SMALL, bg=C_BORDER,
                                   fg=C_TEXT, relief="flat",
                                   padx=10, pady=4, cursor="hand2")
        self.btn_next.pack(side="left", padx=(2,0))
        tk.Label(pag_frm, text=f"({self._page_size} lignes/page)",
                 font=FONT_SMALL, fg=C_MUTED, bg=C_SURFACE).pack(side="left", padx=8)

    # ── Onglet détail ─────────────────────────────────────────────────────

    def _build_detail_tab(self, parent):
        canvas = tk.Canvas(parent, bg=C_BG, highlightthickness=0)
        vsb = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(fill="both", expand=True)

        df = tk.Frame(canvas, bg=C_BG, padx=20, pady=16)
        wid = canvas.create_window((0, 0), window=df, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(wid, width=e.width))
        df.bind("<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # Header
        hdr = tk.Frame(df, bg=C_PRIMARY, pady=14, padx=16)
        hdr.pack(fill="x", pady=(0, 16))
        self.detail_title_var = tk.StringVar(value="Sélectionnez une notice dans le tableau…")
        tk.Label(hdr, textvariable=self.detail_title_var,
                 font=FONT_TITLE, fg="white", bg=C_PRIMARY,
                 wraplength=700, justify="left").pack(anchor="w")

        # Grille
        grid = tk.Frame(df, bg=C_BG)
        grid.pack(fill="x")

        self.detail_vars = {}
        detail_fields = [
            ("N° Inventaire", "N_INVENTAIRE"), ("Fonds",      "CODE_FONDS"),
            ("Cote",          "COTE"),          ("Année",      "ANNEE"),
            ("Auteur",        "NOM_AUTEUR"),    ("Nb pages",   "NB_PAGES"),
            ("Lieu édition",  "NOM_LIEU"),      ("Éditeur",    "NOM_EDITEUR"),
            ("Matière",       "NOM_MATIERE"),
            ("Date ajout",    "DATE_AJOUT"),
        ]
        for i, (label, key) in enumerate(detail_fields):
            row, col = divmod(i, 2)
            colspan = 2 if key in ("NOM_MATIERE", "DATE_AJOUT") else 1
            cell = tk.Frame(grid, bg=C_SURFACE, bd=1, relief="solid",
                            pady=8, padx=12)
            cell.grid(row=row, column=col*2, columnspan=colspan*2,
                      padx=4, pady=4, sticky="ew")
            tk.Label(cell, text=label.upper(), font=FONT_SMALL,
                     fg=C_MUTED, bg=C_SURFACE).pack(anchor="w")
            var = tk.StringVar(value="—")
            self.detail_vars[key] = var
            tk.Label(cell, textvariable=var, font=FONT_BODY,
                     fg=C_TEXT, bg=C_SURFACE,
                     wraplength=320, justify="left").pack(anchor="w", pady=(2,0))

        for c in range(4):
            grid.columnconfigure(c, weight=1)

        # Boutons
        btn_frm = tk.Frame(df, bg=C_BG, pady=12)
        btn_frm.pack(fill="x")
        for txt, cmd, bg in [
            ("✏️  Modifier",   self._edit_notice,   C_PRIMARY),
            ("🗑  Supprimer",  self._delete_notice, C_ACCENT),
        ]:
            tk.Button(btn_frm, text=txt, command=cmd,
                      font=FONT_BODY, bg=bg, fg="white",
                      relief="flat", padx=14, pady=6,
                      cursor="hand2").pack(side="left", padx=(0, 8))

    # ── Onglet Import/Export ──────────────────────────────────────────────

    def _build_io_tab(self, parent):
        outer = tk.Frame(parent, bg=C_BG, padx=20, pady=20)
        outer.pack(fill="both", expand=True)

        # Import
        frm_i = ttk.LabelFrame(outer, text="  📂  Importer vers MySQL  ", padding=12)
        frm_i.pack(fill="x", pady=(0,16))
        for txt, cmd, color in [
            ("📂  Importer BUA (XLS arabe)…",    self._import_bua, C_BUA),
            ("📂  Importer BUF (CSV français)…", self._import_buf, C_BUF),
        ]:
            tk.Button(frm_i, text=txt, command=cmd,
                      font=FONT_BODY, bg=color, fg="white",
                      relief="flat", padx=14, pady=6,
                      cursor="hand2").pack(side="left", padx=(0, 10))
        tk.Label(frm_i,
                 text="⚠️  L'import ignore les numeros d'inventaire deja presents.",
                 font=FONT_SMALL, fg=C_MUTED, bg=C_BG).pack(anchor="w", pady=(10,0))

        # Barre de progression import
        self.prog_var = tk.DoubleVar(value=0)
        self.prog_bar = ttk.Progressbar(frm_i, variable=self.prog_var,
                                         maximum=100, length=400)
        self.prog_bar.pack(fill="x", pady=(8,0))

        # Export
        frm_e = ttk.LabelFrame(outer, text="  💾  Exporter depuis MySQL  ", padding=12)
        frm_e.pack(fill="x", pady=(0,16))
        for txt, cmd, color in [
            ("💾  Export tout (CSV)",  self._export_csv,                C_PRIMARY),
            ("💾  Export BUA",         lambda: self._export_csv("BUA"), C_BUA),
            ("💾  Export BUF",         lambda: self._export_csv("BUF"), C_BUF),
        ]:
            tk.Button(frm_e, text=txt, command=cmd,
                      font=FONT_BODY, bg=color, fg="white",
                      relief="flat", padx=14, pady=6,
                      cursor="hand2").pack(side="left", padx=(0, 8))

        # Log
        frm_log = ttk.LabelFrame(outer, text="  📋  Journal  ", padding=8)
        frm_log.pack(fill="both", expand=True)
        self.log = tk.Text(frm_log, font=FONT_MONO, bg="#0D1B2A",
                            fg="#A8D8A8", height=10, state="disabled",
                            relief="flat", bd=0)
        log_vsb = ttk.Scrollbar(frm_log, orient="vertical", command=self.log.yview)
        self.log.config(yscrollcommand=log_vsb.set)
        log_vsb.pack(side="right", fill="y")
        self.log.pack(fill="both", expand=True)

    # ── Barre de statut ───────────────────────────────────────────────────

    def _build_statusbar(self):
        bar = tk.Frame(self, bg=C_PRIMARY, pady=4)
        bar.pack(fill="x", side="bottom")
        self._status_var = tk.StringVar(value="Prêt — En attente de connexion MySQL.")
        tk.Label(bar, textvariable=self._status_var, font=FONT_SMALL,
                 fg="white", bg=C_PRIMARY, padx=12).pack(side="left")
        tk.Label(bar, text=f"SIGB v{APP_VERSION}  |  MySQL",
                 font=FONT_SMALL, fg=C_GOLD, bg=C_PRIMARY,
                 padx=12).pack(side="right")

    # ══════════════════════════════════════════════════════════════════════
    #  CONNEXION
    # ══════════════════════════════════════════════════════════════════════

    def _auto_connect(self):
        """Tentative de connexion automatique avec les valeurs CONFIG."""
        self._do_connect(DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASSWORD)

    def _open_connect_dialog(self):
        ConnectDialog(self, on_connect=self._do_connect)

    def _do_connect(self, host, port, database, user, password):
        self._status("Connexion a MySQL...")
        self.dao = MySQLDAO(host, port, database, user, password)

        def task():
            try:
                self.dao.connect()
                self.after(0, self._on_connected)
            except Exception as e:
                err = str(e)
                self.after(0, lambda err=err: self._on_connect_error(err))

        threading.Thread(target=task, daemon=True).start()

    def _on_connected(self):
        self.conn_dot.config(fg=C_ONLINE)
        self.conn_label.config(text=f"Connecte : {self.dao.host}/{self.dao.database}")
        self._log(f"✔ Connexion MySQL reussie ({self.dao.host}/{self.dao.database})")
        self._status(f"Connecte a MySQL - {self.dao.database}")
        # Initialiser le schéma si nécessaire
        try:
            created = self.dao.init_schema()
            if created:
                self._log("✔ Schema MySQL cree")
        except Exception as e:
            self._log(f"⚠️ Schéma : {e}")
        self._search()
        self._refresh_stats()

    def _on_connect_error(self, err):
        self.conn_dot.config(fg=C_OFFLINE)
        self.conn_label.config(text="Déconnecté")
        self._log(f"✖ Erreur connexion : {err}")
        self._status("Erreur de connexion MySQL.")
        messagebox.showerror(
            "Connexion MySQL echouee",
            f"Impossible de se connecter :\n\n{err}\n\n"
            "Verifiez le fichier .env ou utilisez Menu -> Base de donnees -> Connecter..."
        )

    def _disconnect(self):
        if self.dao:
            self.dao.disconnect()
        self.conn_dot.config(fg=C_OFFLINE)
        self.conn_label.config(text="Déconnecté")
        self._status("Déconnecté.")
        self._log("Deconnecte de MySQL.")

    def _init_schema(self):
        if not self._check_conn():
            return
        try:
            self.dao.init_schema()
            messagebox.showinfo("Schéma", "Tables créées / déjà existantes.")
        except Exception as e:
            messagebox.showerror("Erreur DDL", str(e))

    def _check_conn(self):
        if not self.dao or not self.dao.is_connected():
            messagebox.showwarning(
                "Non connecté",
                "Vous n'etes pas connecte a MySQL.\n"
                "Menu -> Base de donnees -> Connecter..."
            )
            return False
        return True

    # ══════════════════════════════════════════════════════════════════════
    #  RECHERCHE & AFFICHAGE
    # ══════════════════════════════════════════════════════════════════════

    def _search(self, reset_page=True):
        if not self._check_conn():
            return
        if reset_page:
            self._page = 0
        self._status("Recherche en cours…")

        kw     = self.var_kw.get().strip()
        field  = self.var_field.get()
        fonds  = self.var_fonds.get()
        yr1    = self.var_yr1.get().strip() or None
        yr2    = self.var_yr2.get().strip() or None

        def task():
            try:
                rows, total = self.dao.search(
                    keyword=kw, field=field, fonds=fonds,
                    yr_from=yr1, yr_to=yr2,
                    sort_col=self._sort_col, sort_asc=self._sort_asc,
                    limit=self._page_size, offset=self._page * self._page_size,
                )
                self.after(0, lambda: self._display_results(rows, total))
            except Exception as e:
                err = e
                self.after(0, lambda err=err: self._status(f"Erreur : {err}"))
                self.after(0, lambda err=err: self._log(f"✖ Recherche : {err}"))

        threading.Thread(target=task, daemon=True).start()

    def _display_results(self, rows, total):
        self._rows  = rows
        self._total = total
        self.tree.delete(*self.tree.get_children())

        for i, row in enumerate(rows):
            vals = [str(row.get(c, "") or "") for c in COLUMNS]
            tag  = "odd" if i % 2 == 0 else "even"
            self.tree.insert("", "end", iid=str(i), values=vals, tags=(tag,))

        pages      = max(1, -(-self._total // self._page_size))
        cur_page   = self._page + 1
        self.lbl_page.config(text=f"Page {cur_page} / {pages}")
        self.btn_prev.config(state="normal" if self._page > 0 else "disabled")
        self.btn_next.config(state="normal" if cur_page < pages else "disabled")
        self.lbl_count.config(
            text=f"{self._total:,} notices".replace(",", " "))
        self.stats_vars["shown"].set(str(len(rows)))
        self._status(f"{self._total:,} notices trouvées.".replace(",", " "))

    def _next_page(self):
        self._page += 1
        self._search(reset_page=False)

    def _prev_page(self):
        if self._page > 0:
            self._page -= 1
            self._search(reset_page=False)

    def _refresh_stats(self):
        if not self._check_conn():
            return
        def task():
            try:
                st = self.dao.get_stats()
                self.after(0, lambda: [
                    self.stats_vars[k].set(f"{v:,}".replace(",", " "))
                    for k, v in st.items() if k in self.stats_vars
                ])
            except Exception:
                pass
        threading.Thread(target=task, daemon=True).start()

    # ── Tri ───────────────────────────────────────────────────────────────

    def _sort_by(self, col):
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True
        arrow = " ↑" if self._sort_asc else " ↓"
        for c in COLUMNS:
            self.tree.heading(c, text=COL_LABELS[c] + (arrow if c == col else ""))
        self._search()

    # ── Sélection ─────────────────────────────────────────────────────────

    def _on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        row = self._rows[idx]
        self._sel_data = row
        self._sel_id   = None  # on ira le chercher si besoin

        self.detail_title_var.set(row.get("TITRE") or "—")
        for key, var in self.detail_vars.items():
            val = row.get(key, "") or "—"
            var.set(str(val))

    def _on_double_click(self, event=None):
        self.notebook.select(1)

    def _get_id_notice(self):
        """Récupère l'ID_NOTICE de la ligne sélectionnée."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Aucune sélection",
                                "Sélectionnez d'abord une notice dans le tableau.")
            return None
        row = self._rows[int(sel[0])]
        if row.get("ID_NOTICE"):
            return row["ID_NOTICE"]
        inv = str(row.get("N_INVENTAIRE", "")).strip()
        try:
            cur = self.dao.conn.cursor()
            cur.execute("""
                SELECT id_notice
                FROM exemplaire
                WHERE num_inventaire=%s
                LIMIT 1
            """, (inv,))
            r = cur.fetchone()
            cur.close()
            return r[0] if r else None
        except Exception:
            return None

    # ══════════════════════════════════════════════════════════════════════
    #  CRUD MYSQL
    # ══════════════════════════════════════════════════════════════════════

    # ── Ajouter ──────────────────────────────────────────────────────────

    def _add_notice(self):
        if not self._check_conn():
            return
        def on_save(data):
            try:
                self.dao.add_notice(data)
                self._log(f"✔ Ajout MySQL — Inv. {data['N_INVENTAIRE']} | {data['TITRE'][:50]}")
                self._status("Notice ajoutee dans MySQL.")
                self._search()
                self._refresh_stats()
            except Exception as e:
                messagebox.showerror("Erreur MySQL", f"INSERT echoue :\n{e}")
                self._log(f"✖ Ajout échoué : {e}")
        NoticeDialog(self, title="➕  Nouvelle notice", on_save=on_save)

    # ── Modifier ─────────────────────────────────────────────────────────

    def _edit_notice(self):
        if not self._check_conn():
            return
        id_notice = self._get_id_notice()
        if not id_notice:
            return
        # Recharger la notice complete depuis MySQL
        full = self.dao.get_notice(id_notice)
        if not full:
            messagebox.showerror("Erreur", "Notice introuvable en base.")
            return
        def on_save(data):
            try:
                self.dao.update_notice(id_notice, data)
                self._log(f"✔ Modif MySQL — Inv. {data['N_INVENTAIRE']} | {data['TITRE'][:50]}")
                self._status("Notice modifiee dans MySQL.")
                self._search()
            except Exception as e:
                messagebox.showerror("Erreur MySQL", f"UPDATE echoue :\n{e}")
                self._log(f"✖ Modif échouée : {e}")
        NoticeDialog(self, title="✏️  Modifier la notice",
                     data=full, on_save=on_save)

    # ── Supprimer ─────────────────────────────────────────────────────────

    def _delete_notice(self):
        if not self._check_conn():
            return
        id_notice = self._get_id_notice()
        if not id_notice:
            return
        titre = (self._sel_data.get("TITRE") or "")[:60]
        inv   = self._sel_data.get("N_INVENTAIRE", "?")
        if not messagebox.askyesno(
                "Confirmer la suppression",
                f"Supprimer definitivement de la base MySQL :\n\n"
                f"«{titre}»\n(N° inventaire : {inv}) ?",
                icon="warning"):
            return
        try:
            self.dao.delete_notice(id_notice)
            self._log(f"🗑 Supprime MySQL — ID {id_notice} | Inv. {inv}")
            self._status(f"Notice supprimée (Inv. {inv}).")
            self._search()
            self._refresh_stats()
        except Exception as e:
            messagebox.showerror("Erreur MySQL", f"DELETE echoue :\n{e}")
            self._log(f"✖ Suppression échouée : {e}")

    # ══════════════════════════════════════════════════════════════════════
    #  IMPORT
    # ══════════════════════════════════════════════════════════════════════

    def _import_menu(self):
        m = tk.Menu(self, tearoff=0, bg=C_SURFACE, fg=C_TEXT, font=FONT_BODY)
        m.add_command(label="📂  Importer BUA (XLS)…", command=self._import_bua)
        m.add_command(label="📂  Importer BUF (CSV)…", command=self._import_buf)
        m.post(self.winfo_pointerx(), self.winfo_pointery())

    def _import_bua(self):
        if not self._check_conn():
            return
        path = filedialog.askopenfilename(
            title="Choisir BUA (XLS)",
            filetypes=[("Excel 97-2003", "*.xls"), ("Tous", "*.*")])
        if not path:
            return
        self._run_import(path, "BUA", "xls")

    def _import_buf(self):
        if not self._check_conn():
            return
        path = filedialog.askopenfilename(
            title="Choisir BUF (CSV)",
            filetypes=[("CSV", "*.csv"), ("Tous", "*.*")])
        if not path:
            return
        self._run_import(path, "BUF", "csv")

    def _run_import(self, path, fonds, fmt):
        self._status(f"Import {fonds} en cours…")
        self._log(f"Import {fonds} depuis {os.path.basename(path)}…")
        self.prog_var.set(0)

        def task():
            try:
                if not PANDAS_OK:
                    raise RuntimeError(
                        "pandas non installé. Installez les dépendances : pip install pandas xlrd openpyxl"
                    )
                import pandas as pd
                if fmt == "xls":
                    df = pd.read_excel(path, engine="xlrd", dtype=str)
                    df.columns = ["Cote","Titre","Auteur","Lieu","Edition",
                                  "Annee","Nb_pages","Matiere","Inventaire"]
                else:
                    df = pd.read_csv(path, sep=";", encoding="latin1",
                                      header=None, dtype=str)
                    df.columns = ["Cote","Titre","Auteur","Lieu","Edition",
                                  "Annee","Nb_pages","Matiere","Inventaire"]
                df = df.fillna("")
                rows = df.to_dict(orient="records")

                def progress(done, total):
                    pct = done / total * 100
                    self.after(0, lambda: self.prog_var.set(pct))

                ok = self.dao.bulk_insert(rows, fonds, progress_cb=progress)
                self.after(0, lambda: self.prog_var.set(100))
                self.after(0, lambda: self._log(
                    f"✔ Import termine : {ok} notices inserees dans MySQL ({fonds})"))
                self.after(0, lambda: self._status(
                    f"Import {fonds} terminé : {ok} notices."))
                self.after(0, self._search)
                self.after(0, self._refresh_stats)
            except Exception as e:
                err = str(e)
                self.after(0, lambda err=err: self._log(f"✖ Import échoué : {err}"))
                self.after(0, lambda err=err: messagebox.showerror("Erreur import", err))

        threading.Thread(target=task, daemon=True).start()

    # ══════════════════════════════════════════════════════════════════════
    #  EXPORT CSV
    # ══════════════════════════════════════════════════════════════════════

    def _export_csv(self, fonds=None):
        if not self._check_conn():
            return
        path = filedialog.asksaveasfilename(
            title="Exporter CSV",
            defaultextension=".csv",
            initialfile=f"sigb_{fonds or 'all'}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            filetypes=[("CSV", "*.csv")])
        if not path:
            return
        self._status("Export CSV en cours…")

        def task():
            try:
                rows, _ = self.dao.search(
                    fonds=fonds or "TOUS",
                    limit=99999, offset=0,
                )
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.DictWriter(f, fieldnames=COLUMNS,
                                            delimiter=";", extrasaction="ignore")
                    writer.writeheader()
                    writer.writerows(rows)
                self.after(0, lambda: self._log(
                    f"✔ Export CSV → {os.path.basename(path)} ({len(rows)} lignes)"))
                self.after(0, lambda: self._status("Export CSV terminé."))
                self.after(0, lambda: messagebox.showinfo(
                    "Export réussi", f"{len(rows)} notices exportées\n→ {path}"))
            except Exception as e:
                err = str(e)
                self.after(0, lambda err=err: messagebox.showerror("Erreur export", err))

        threading.Thread(target=task, daemon=True).start()

    # ══════════════════════════════════════════════════════════════════════
    #  UTILITAIRES
    # ══════════════════════════════════════════════════════════════════════

    def _reset(self):
        self.var_kw.set("")
        self.var_field.set("TOUS")
        self.var_fonds.set("TOUS")
        self.var_yr1.set("")
        self.var_yr2.set("")
        self._page = 0
        self._search()

    def _status(self, msg):
        self._status_var.set(msg)

    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log.config(state="normal")
        self.log.insert("end", f"[{ts}]  {msg}\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def _about(self):
        messagebox.showinfo(
            f"À propos — SIGB v{APP_VERSION}",
            f"SIGB — Système Intégré de Gestion Bibliothécaire\n"
            f"Version {APP_VERSION}\n\n"
            f"Base de donnees : MySQL\n"
            f"Driver : mysql-connector-python\n"
            f"Configuration : fichier .env\n\n"
            f"Fonctions : Recherche · Ajout · Modification\n"
            f"            Suppression · Import bulk · Export CSV\n\n"
            f"Dépendances :\n"
            f"  pip install mysql-connector-python python-dotenv pandas xlrd openpyxl"
        )


# ══════════════════════════════════════════════════════════════════════════════
#  POINT D'ENTRÉE
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = SIGBApp()
    app.mainloop()
