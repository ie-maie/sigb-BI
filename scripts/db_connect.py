import os
from pathlib import Path

import mysql.connector
from mysql.connector import Error
from dotenv import load_dotenv

ROOT_DIR = Path(__file__).resolve().parents[1]
ENV_PATH = ROOT_DIR / ".env"

load_dotenv(ENV_PATH)


def get_connection():
    host = os.getenv("DB_HOST", "localhost")
    port = int(os.getenv("DB_PORT", 3306))
    database = os.getenv("DB_NAME", "sigb")
    user = os.getenv("DB_USER", "sigb_user")
    password = os.getenv("DB_PASSWORD", "")

    try:
        conn = mysql.connector.connect(
            host=host,
            port=port,
            database=database,
            user=user,
            password=password,
            charset="utf8mb4",
            connection_timeout=10,
        )
        if not conn.is_connected():
            raise RuntimeError("La connexion MySQL a ete creee mais elle n'est pas active.")
        return conn
    except Error as exc:
        raise RuntimeError(
            "Impossible de se connecter a MySQL. Verifie le fichier .env "
            f"(DB_HOST={host}, DB_PORT={port}, DB_NAME={database}, DB_USER={user}). "
            f"Erreur MySQL: {exc}"
        ) from exc
