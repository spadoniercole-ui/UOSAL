import psycopg2
DATABASE_URL = 'postgresql://postgres:Mges/0905%211@db.cjfkqhwoxbjeznarvazm.supabase.co:5432/postgres'

def test_supabase_connection():
    try:
        connection = psycopg2.connect(DATABASE_URL)
        print("Connessione a Supabase (PostgreSQL) riuscita con successo! 🚀")
        
        with connection.cursor() as cursor:
            cursor.execute("SELECT version();")
            db_version = cursor.fetchone()
            print(f"\nVersione Database: {db_version[0]}")
            
            cursor.execute("SELECT table_name FROM information_schema.tables WHERE table_schema = 'public';")
            tabelle = cursor.fetchall()
            
            print("\nTabelle presenti nel database:")
            for t in tabelle:
                print(f" - {t[0]}")
                
    except Exception as e:
        print(f"Errore di connessione a Supabase: {e}")
    finally:
        if 'connection' in locals() and connection.closed == 0:
            connection.close()

if __name__ == "__main__":
    test_supabase_connection()
