#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from db_connect import get_connection


def main():
    conn = get_connection()
    print("Connexion MySQL reussie.")
    print("Base active :", conn.database)
    conn.close()


if __name__ == "__main__":
    main()
