import os
import pymysql

# Coordinate di Railway fornite
HOST = "viaduct.proxy.rlwy.net"  # Inserisci qui l'host pubblico di Railway
PORT = 12345                     # Inserisci qui la porta a 5 cifre del proxy

USER = "root"
PASSWORD = "nEVTHKFiCxqaNRqQuvNHXQOikjqXsotB"
DATABASE = "railway"

def test_connection():
    try:
        connection = pymysql.connect(
            host=mysql.railway.internal,
            port=3306,
            user=root,
            password=nEVTHKFiCxqaNRqQuvNHXQOikjqXsotB,
            database=railway,
            cursorclass=pymysql.cursors.DictCursor
        )
        print("Connessione a MySQL su Railway riuscita con successo! 🚀")
        
        with connection.cursor() as cursor:
            # Test di lettura dei parametri di sistema
            cursor.execute("SELECT chiave, valore, descrizione FROM parametri_sistema;")
            parametri = cursor.fetchall()
            
            print("\nParametri di sistema configurati nel database:")
            for p in parametri:
                print(f" - [{p['chiave']}] => {p['valore']} ({p['descrizione']})")
                
    except Exception as e:
        print(f"Errore di connessione al database: {e}")
    finally:
        if 'connection' in locals() and connection.open:
            connection.close()

if __name__ == "__main__":
    test_connection()
