import pymysql

# SOSTITUISCI QUESTI DUE VALORI CON QUELLI DEL TCP PROXY DI RAILWAY:
HOST = "viaduct.proxy.rlwy.net"  # Metti qui il tuo host pubblico di Railway
PORT = 12345                     # Metti qui la porta a 5 cifre del proxy

USER = "root"
PASSWORD = "nEVTHKFiCxqaNRqQuvNHXQOikjqXsotB"
DATABASE = "railway"

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