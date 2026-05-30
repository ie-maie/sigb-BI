import mysql.connector
import os
from dotenv import load_dotenv

load_dotenv()

def get_connection():
    return mysql.connector.connect(
        host=os.getenv("DB_HOST", "localhost"),
        port=int(os.getenv("DB_PORT", 3306)),
        database=os.getenv("DB_NAME", "sigb"),
        user=os.getenv("DB_USER", "sigb_user"),
        password=os.getenv("DB_PASSWORD", ""),
        charset="utf8mb4"
    )
