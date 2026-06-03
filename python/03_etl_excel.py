#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETL - bua.xls (fonds arabophone ~10179 notices)
Feuil1: Cote|Titre|Auteur|Lieu|Edition|Annee|Nb pages|Matiere|Inventaire
Feuil2: Cote|Inventaire
"""

import re
import unicodedata
import xlrd
from db_connect import ROOT_DIR, get_connection

XLS_PATH  = ROOT_DIR / "data" / "bua.xls"
LANG_CODE = "ara"

_cache_editeur = {}
_cache_auteur  = {}
_cache_matiere = {}
_cache_classif = {}


def clean(val):
    if val is None:
        return ""
    s = str(val).strip()
    if s.startswith('"'):
        s = s[1:]
    s = "".join(c for c in s if unicodedata.category(c)[0] != "C")
    return s.strip()


def safe_int(val, default=None):
    try:
        s = str(val).strip()
        return int(float(s)) if s else default
    except (ValueError, TypeError):
        return default


def parse_auteur_arabe(raw):
    raw = clean(raw)
    if not raw:
        return None, None
    if "," in raw:
        parts = raw.split(",", 1)
        return clean(parts[0]), clean(parts[1])
    return raw, None


def upsert_editeur(cursor, nom, ville):
    nom   = clean(nom)
    ville = clean(ville) or None
    if not nom:
        return None
    key = (nom.lower(), (ville or "").lower())
    if key in _cache_editeur:
        return _cache_editeur[key]
    cursor.execute(
        "INSERT IGNORE INTO editeur (nom_editeur, ville) VALUES (%s, %s)", (nom, ville))
    cursor.execute(
        "SELECT id_editeur FROM editeur WHERE nom_editeur=%s AND COALESCE(ville,'')=COALESCE(%s,'')",
        (nom, ville))
    row = cursor.fetchone()
    if row:
        _cache_editeur[key] = row[0]
        return row[0]
    return None


def upsert_auteur(cursor, nom, prenom):
    if not nom:
        return None
    prenom = prenom or None
    key = (nom.lower(), (prenom or "").lower())
    if key in _cache_auteur:
        return _cache_auteur[key]
    cursor.execute(
        "INSERT IGNORE INTO auteur (nom, prenom) VALUES (%s, %s)", (nom, prenom))
    cursor.execute(
        "SELECT id_auteur FROM auteur WHERE nom=%s AND COALESCE(prenom,'')=COALESCE(%s,'')",
        (nom, prenom))
    row = cursor.fetchone()
    if row:
        _cache_auteur[key] = row[0]
        return row[0]
    return None


def upsert_matiere(cursor, libelle):
    libelle = clean(libelle)
    if not libelle:
        return None
    key = libelle.lower()
    if key in _cache_matiere:
        return _cache_matiere[key]
    cursor.execute("INSERT IGNORE INTO matiere (libelle) VALUES (%s)", (libelle,))
    cursor.execute("SELECT id_matiere FROM matiere WHERE libelle=%s", (libelle,))
    row = cursor.fetchone()
    if row:
        _cache_matiere[key] = row[0]
        return row[0]
    return None


def upsert_classif(cursor, cote):
    cote = clean(cote)
    if not cote:
        return None
    if cote in _cache_classif:
        return _cache_classif[cote]
    cursor.execute("INSERT IGNORE INTO classification (cote) VALUES (%s)", (cote,))
    cursor.execute("SELECT id_classification FROM classification WHERE cote=%s", (cote,))
    row = cursor.fetchone()
    if row:
        _cache_classif[cote] = row[0]
        return row[0]
    return None


def load_feuil1(cursor, wb, id_langue):
    sh = wb.sheet_by_index(0)
    inserted = skipped = errors = 0

    for r in range(1, sh.nrows):
        try:
            row = [sh.cell_value(r, c) for c in range(sh.ncols)]
            while len(row) < 9:
                row.append("")

            cote_raw, titre_raw, auteur_raw, lieu_raw, editeur_raw, \
                annee_raw, pages_raw, matiere_raw, inv_raw = row[:9]

            titre = clean(titre_raw)
            if not titre:
                skipped += 1
                continue

            id_editeur = upsert_editeur(cursor, editeur_raw, lieu_raw)
            id_classif = upsert_classif(cursor, cote_raw)
            annee      = safe_int(annee_raw)
            pages      = safe_int(pages_raw)

            cursor.execute("""
                INSERT INTO notice
                    (titre, annee_pub, nb_pages, id_editeur, id_langue, id_classification)
                VALUES (%s,%s,%s,%s,%s,%s)
            """, (titre, annee, pages, id_editeur, id_langue, id_classif))
            id_notice = cursor.lastrowid

            for ar in re.split(r";\s*", clean(str(auteur_raw))):
                nom, prenom = parse_auteur_arabe(ar)
                id_aut = upsert_auteur(cursor, nom, prenom)
                if id_aut:
                    cursor.execute(
                        "INSERT IGNORE INTO notice_auteur (id_notice, id_auteur) VALUES (%s,%s)",
                        (id_notice, id_aut))

            for m in re.split(r"\s*[.;]\s*", clean(str(matiere_raw))):
                id_mat = upsert_matiere(cursor, m.strip())
                if id_mat:
                    cursor.execute(
                        "INSERT IGNORE INTO notice_matiere (id_notice, id_matiere) VALUES (%s,%s)",
                        (id_notice, id_mat))

            inv = clean(str(inv_raw))
            if inv and inv != "0":
                inv_int = safe_int(inv)
                inv_str = str(inv_int) if inv_int else None
                if inv_str:
                    cursor.execute("""
                        INSERT IGNORE INTO exemplaire
                            (num_inventaire, id_notice, cote_exemplaire)
                        VALUES (%s,%s,%s)
                    """, (inv_str, id_notice, clean(str(cote_raw))))

            inserted += 1
            if inserted % 1000 == 0:
                print(f"  {inserted} notices inserees...")

        except Exception as e:
            errors += 1
            print(f"  Ligne {r}: {e}")

    print(f"Feuil1: {inserted} notices, {skipped} ignorees, {errors} erreurs")


def load_feuil2(cursor, wb):
    sh = wb.sheet_by_index(1)
    inserted = skipped = 0

    for r in range(1, sh.nrows):
        cote_raw = clean(str(sh.cell_value(r, 0)))
        inv_raw  = sh.cell_value(r, 1)

        if not cote_raw:
            skipped += 1
            continue

        inv     = safe_int(inv_raw)
        inv_str = str(inv) if inv else None

        # Étape 1 : trouver la classification
        cursor.execute(
            "SELECT id_classification FROM classification WHERE cote=%s LIMIT 1",
            (cote_raw,))
        res = cursor.fetchone()

        if not res:
            skipped += 1
            continue

        # Étape 2 : trouver la notice liée
        cursor.execute(
            "SELECT id_notice FROM notice WHERE id_classification=%s LIMIT 1",
            (res[0],))
        row = cursor.fetchone()

        if row and inv_str:
            cursor.execute("""
                INSERT IGNORE INTO exemplaire (num_inventaire, id_notice, cote_exemplaire)
                VALUES (%s,%s,%s)
            """, (inv_str, row[0], cote_raw))
            inserted += 1
        else:
            skipped += 1

    print(f"Feuil2: {inserted} exemplaires ajoutes, {skipped} ignores")


def load_excel():
    if not XLS_PATH.exists():
        raise FileNotFoundError(f"Fichier Excel introuvable : {XLS_PATH}")

    conn   = get_connection()
    cursor = conn.cursor(buffered=True)

    cursor.execute("SELECT id_langue FROM langue WHERE code_langue=%s", (LANG_CODE,))
    row_langue = cursor.fetchone()
    if not row_langue:
        raise RuntimeError("La langue 'ara' est absente de la table langue.")
    id_langue = row_langue[0]

    wb = xlrd.open_workbook(str(XLS_PATH))

    print("Chargement Feuil1...")
    load_feuil1(cursor, wb, id_langue)
    conn.commit()

    print("Chargement Feuil2...")
    load_feuil2(cursor, wb)
    conn.commit()

    cursor.close()
    conn.close()


if __name__ == "__main__":
    print("Chargement bua.xls...")
    load_excel()
    print("Termine.")
