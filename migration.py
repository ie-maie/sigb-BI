"""
Migration script for SIGB BI database
Connects to remote MySQL VM and manages schema migration
"""


import os
import sys
from pathlib import Path
from dotenv import load_dotenv
import mysql.connector
from mysql.connector import Error

# Load environment variables
load_dotenv()

# Database configuration from .env
DB_CONFIG = {
    'host': os.getenv('DB_HOST', '84.8.219.151'),
    'user': os.getenv('DB_USER', 'sigb_user'),
    'password': os.getenv('DB_PASSWORD', 'Sigb2024!'),
    'database': os.getenv('DB_NAME', 'sigb'),
    'raise_on_warnings': False,
}

def connect_to_db():
    """Connect to remote MySQL database on VM"""
    try:
        connection = mysql.connector.connect(**DB_CONFIG)
        if connection.is_connected():
            db_info = connection.get_server_info()
            print(f"✅ Connected to MySQL Server version {db_info}")
            print(f"   Host: {DB_CONFIG['host']}")
            print(f"   Database: {DB_CONFIG['database']}")
            return connection
    except Error as e:
        print(f"❌ Error connecting to MySQL: {e}")
        sys.exit(1)

def execute_migration_sql(connection):
    """Execute the migration SQL script"""
    cursor = connection.cursor()
    
    # Read migration SQL file
    sql_file = Path('sql/00_migration_clean.sql')
    if not sql_file.exists():
        print(f"❌ Migration SQL file not found: {sql_file}")
        return False
    
    print(f"\n📋 Reading migration SQL from: {sql_file}")
    
    try:
        with open(sql_file, 'r', encoding='utf-8') as f:
            sql_script = f.read()
        
        print("📊 Executing SQL script (safe splitter: statement terminates only on \n;\n)" )
        print("=" * 80)

        # Robust splitter for this repo SQL: split only when we see a statement terminator
        # on its own line: "\n;\n" (our migration file uses this style).
        statements = []
        buff = []
        for line in sql_script.splitlines(True):
            buff.append(line)
            if line.strip() == ';':
                stmt = ''.join(buff).strip()
                if stmt:
                    statements.append(stmt)
                buff = []
        # If anything remains, try to execute it as well
        tail = ''.join(buff).strip()
        if tail:
            statements.append(tail)

        print(f"📊 Found {len(statements)} SQL statements to execute")

        for i, statement in enumerate(statements, 1):
            try:
                cursor.execute(statement)
                print(f"✅ [{i}/{len(statements)}] {statement[:60]}...")
            except Error as e:
                print(f"⚠️  [{i}/{len(statements)}] {statement[:60]}...")
                print(f"   Warning: {e}")

        connection.commit()



        print("=" * 80)
        print("✅ Migration SQL executed successfully")
        return True
        
    except Exception as e:
        print(f"❌ Error executing migration SQL: {e}")
        connection.rollback()
        return False
    finally:
        cursor.close()

def verify_schema(connection):
    """Verify that all required tables exist"""
    cursor = connection.cursor()
    
    required_tables = [
        'langue',
        'editeur',
        'auteur',
        'classification',
        'matiere',
        'notice',
        'notice_auteur',
        'notice_matiere',
        'exemplaire'
    ]
    
    print("\n🔍 Verifying schema...")
    print("-" * 80)
    
    try:
        cursor.execute("SHOW TABLES;")
        existing_tables = [table[0] for table in cursor.fetchall()]
        
        all_exist = True
        for table in required_tables:
            if table in existing_tables:
                print(f"✅ Table '{table}' exists")
            else:
                print(f"❌ Table '{table}' MISSING")
                all_exist = False
        
        print("-" * 80)
        if all_exist:
            print(f"✅ All {len(required_tables)} required tables exist")
            return True
        else:
            print("❌ Some tables are missing")
            return False
            
    except Error as e:
        print(f"❌ Error verifying schema: {e}")
        return False
    finally:
        cursor.close()

def get_table_stats(connection):
    """Get row counts for all tables"""
    cursor = connection.cursor()
    
    print("\n📊 Table Statistics:")
    print("-" * 80)
    
    try:
        cursor.execute("""
            SELECT TABLE_NAME, TABLE_ROWS 
            FROM INFORMATION_SCHEMA.TABLES 
            WHERE TABLE_SCHEMA = %s
            ORDER BY TABLE_ROWS DESC
        """, (DB_CONFIG['database'],))
        
        total_rows = 0
        for table_name, row_count in cursor.fetchall():
            print(f"  {table_name:20} → {row_count:>10,} rows")
            total_rows += row_count
        
        print("-" * 80)
        print(f"  {'TOTAL':20} → {total_rows:>10,} rows")
        
    except Error as e:
        print(f"⚠️  Could not retrieve table statistics: {e}")
    finally:
        cursor.close()

def main():
    """Main migration process"""
    print("\n" + "=" * 80)
    print("🔄 SIGB BI Database Migration")
    print("=" * 80)
    
    # Step 1: Connect
    print("\n[1/4] Connecting to remote MySQL VM...")
    connection = connect_to_db()
    
    try:
        # Step 2: Execute migration SQL
        print("\n[2/4] Executing migration SQL...")
        if not execute_migration_sql(connection):
            print("⚠️  Migration SQL had issues but continuing...")
        
        # Step 3: Verify schema
        print("\n[3/4] Verifying schema creation...")
        if not verify_schema(connection):
            print("⚠️  Some tables missing - schema creation may have failed")
        
        # Step 4: Show statistics
        print("\n[4/4] Retrieving table statistics...")
        get_table_stats(connection)
        
        print("\n" + "=" * 80)
        print("✅ Migration process completed")
        print("=" * 80)
        print("\n📝 Next steps:")
        print("   1. Run etl.ipynb to generate CSV files")
        print("   2. Run load_etl_to_mysql.py to load data")
        print("=" * 80 + "\n")
        
    finally:
        if connection.is_connected():
            connection.close()
            print("🔌 Database connection closed")

if __name__ == '__main__':
    main()
