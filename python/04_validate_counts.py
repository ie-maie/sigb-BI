#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from db_connect import get_connection


TABLES = [
    "notice",
    "auteur",
    "editeur",
    "matiere",
    "classification",
    "exemplaire",
]


def main():
    conn = get_connection()
    cursor = conn.cursor()

    print("Controle des volumes :")
    for table in TABLES:
        cursor.execute(f"SELECT COUNT(*) FROM {table}")
        total = cursor.fetchone()[0]
        print(f"- {table}: {total}")

    cursor.close()
    conn.close()


if __name__ == "__main__":
    main()
