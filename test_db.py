import os
import pymysql
from dotenv import load_dotenv

# Carica le variabili d'ambiente (o usa direttamente i dati di Railway)
# Se sei in locale, usa l'host pubblico di Railway (es. viaduct.proxy.rlwy.net e la porta dedicata)
# Se lo script gira su Railway, puoi usare direttamente le variabili d'ambiente fornite.

HOST = os.getenv("MYSQLHOST", "mysql.railway.internal")
PORT = int(os.getenv("MYSQLPORT", 3306))
USER = os.getenv("MYSQLUSER", "root")
PASSWORD = os.getenv("MYSQLPASSWORD", "nEVTHKFiCxqaNRqQuvNHXQOikjqXsotB")
DATABASE = os.getenv("MYSQLDATABASE", "railway")

def test_connection():
    try:
        connection = pymysql.connect(
            host=HOST,
            port=PORT,
            user=USER,
            password=PASSWORD,
            database=DATABASE,
            cursorclass=pymysql.cursors.DictCursor
        )
        print("Connessione a MySQL su Railway riuscita con successo! 🚀")
        
        with connection.cursor() as cursor:
            # Test di lettura dei parametri di sistema
            cursor.execute("SELECT chiave, valore, descrizione FROM parametri_sistema;")
            parametri = cursor.fetchall()
            
            print("\nParametri di sistema configurati nel database:")
            for p in parametri:
                print(f" - [{p['chiave']}] => {p['valore']} ({p['descrizione'])")
                
    except Exception as e:
        print(f"Errore di connessione al database: {e}")
    finally:
        if 'connection' in locals() and connection.open:
            connection.close()

if __name__ == "__main__":
    test_connection()