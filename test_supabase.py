import os
from supabase import create_client, Client

# Le credenziali dirette del tuo progetto Supabase
SUPABASE_URL = "https://cjfkqhwoxbjeznarvazm.supabase.co"
# Inserirai la tua chiave anon o service_role che trovi nelle impostazioni API di Supabase
SUPABASE_KEY = "sb_publishable_pZ3MqT04XNr6C78tHfCBBg_mVp5wfgS"

def test_supabase_http():
    try:
        supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
        print("Connessione via HTTP/REST a Supabase avviata con successo! 🚀")
        
        # Facciamo una query di test sulla tabella parametri_sistema che abbiamo creato
        response = supabase.table("parametri_sistema").select("*").execute()
        
        print("\nParametri letti con successo dal database:")
        for row in response.data:
            print(f" - [{row['chiave']}] => {row['valore']}")
                
    except Exception as e:
        print(f"Errore di comunicazione con Supabase: {e}")

if __name__ == "__main__":
    test_supabase_http()
