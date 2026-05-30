#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETL - buf.csv (fonds francophone ~15190 notices)
Colonnes: Cote;Titre;Auteur;Lieu;Editeur;Annee;Nb_pages;Matiere;Inventaire
"""

import csv
import io
import re
import unicodedata
from pathlib import Path
from db_connect import get_connection

CSV_PATH  = Path("data/buf.csv")
ENCODING  = "latin-1"
SEP       = ";"
LANG_CODE = "fre"

_cache_editeur = {}
_cache_auteur  = {}
_cache_matiere = {}
_cache_classif = {}


def clean(s):
    if not s:
        return ""
    s = str(s).strip()
    s = "".join(c for c in s if unicodedata.category(c)[0] != "C")
    return s


def split_titre_parallele(titre):
    if " = " in titre:
        parts = titre.split(" = ", 1)
        return clean(parts[0]), clean(parts[1])
    return clean(titre), None


def parse_auteur(raw):
    raw = clean(raw)
    if not raw:
        return None, None
    if any(c in raw for c in ["/", "&", "+"]):
        return raw, None
    tokens = raw.split()
    if not tokens:
        return None, None
    nom_parts = []
    prenom_parts = []
    frontier_found = False
    for tok in tokens:
        if not frontier_found and tok.replace("-", "").replace("'", "").isupper():
            nom_parts.append(tok)
        else:
            frontier_found = True
            prenom_parts.append(tok)
    nom    = " ".join(nom_parts).strip()
    prenom = " ".join(prenom_parts).strip()
    if not nom:
        nom = raw
    return nom or None, prenom or None


def safe_int(val, default=None):
    try:
        return int(float(str(val).strip())) if str(val).strip() else default
    except (ValueError, AttributeError):
        return default


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


def load_csv():
    conn   = get_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT id_langue FROM langue WHERE code_langue=%s", (LANG_CODE,))
    id_langue = cursor.fetchone()[0]

    raw  = CSV_PATH.read_bytes()
    text = raw.decode(ENCODING)
    reader = csv.reader(io.StringIO(text), delimiter=SEP)
    rows   = list(reader)

    inserted = 0
    skipped  = 0
    errors   = 0

    for i, row in enumerate(rows):
        try:
            while len(row) < 9:
                row.append("")
            cote, titre_raw, auteur_raw, lieu, editeur_raw, \
                annee_raw, pages_raw, matiere_raw, inventaire_raw = row[:9]

            titre_raw = clean(titre_raw)
            if not titre_raw:
                skipped += 1
                continue

            titre, titre_parallele = split_titre_parallele(titre_raw)
            id_editeur = upsert_editeur(cursor, editeur_raw, lieu)
            id_classif = upsert_classif(cursor, cote)
            annee      = safe_int(annee_raw)
            pages      = safe_int(pages_raw)

            cursor.execute("""
                INSERT INTO notice
                    (titre, titre_parallele, annee_pub, nb_pages,
                     id_editeur, id_langue, id_classification)
                VALUES (%s,%s,%s,%s,%s,%s,%s)
            """, (titre, titre_parallele, annee, pages, id_editeur, id_langue, id_classif))
            id_notice = cursor.lastrowid

            auteurs_raw = re.split(r"\s[./]\s|\s/\s", clean(auteur_raw))
            for ar in auteurs_raw:
                nom, prenom = parse_auteur(ar)
                id_aut = upsert_auteur(cursor, nom, prenom)
                if id_aut:
                    cursor.execute(
                        "INSERT IGNORE INTO notice_auteur (id_notice, id_auteur) VALUES (%s,%s)",
                        (id_notice, id_aut))

            for m in re.split(r"\s\.\s|;", clean(matiere_raw)):
                id_mat = upsert_matiere(cursor, m.strip())
                if id_mat:
                    cursor.execute(
                        "INSERT IGNORE INTO notice_matiere (id_notice, id_matiere) VALUES (%s,%s)",
                        (id_notice, id_mat))

            inv = clean(inventaire_raw)
            if inv:
                cursor.execute("""
                    INSERT IGNORE INTO exemplaire
                        (num_inventaire, id_notice, cote_exemplaire)
                    VALUES (%s,%s,%s)
                """, (inv, id_notice, clean(cote)))

            inserted += 1
            if inserted % 1000 == 0:
                conn.commit()
                print(f"  {inserted} notices inserees...")

        except Exception as e:
            errors += 1
            print(f"  Ligne {i}: {e}")

    conn.commit()
    cursor.close()
    conn.close()
    print(f"\nCSV charge : {inserted} notices, {skipped} ignorees, {errors} erreurs")


if __name__ == "__main__":
    print("Chargement buf.csv...")
    load_csv()
