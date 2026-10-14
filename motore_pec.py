import re
from supabase import create_client, Client

# Credenziali Supabase confermate
SUPABASE_URL = "https://cjfkqhwoxbjeznarvazm.supabase.co"
SUPABASE_KEY = "INSERISCI_QUI_LA_TUA_API_KEY" # Sostituisci con la tua chiave effettiva

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

def carica_parametri():
    """Carica le configurazioni e le regex direttamente dalla tabella parametrica su Supabase"""
    response = supabase.table("parametri_sistema").select("chiave, valore").execute()
    return {row["chiave"]: row["valore"] for row in response.data}

def analizza_e_smista_pec(mittente: str, oggetto: str, corpo: str):
    """Motore logico per il parsing e il matching della PEC"""
    print(f"\n--- Analisi nuova PEC da: {mittente} ---")
    print(f"Oggetto: {oggetto}")
    
    # 1. Carichiamo i parametri dinamici dal database
    parametri = carica_parametri()
    soglia_auto = int(parametri.get("soglia_match_automatico", 85))
    regex_pattern = parametri.get("regex_ruolo_principale", r'/(?:R\.?G\.?|RG|RUOLO)\s*(?:N\.?|NUM\.?\s*)?(\d{1,6})[\/\-](\d{2,4})/ui')
    
    # Pulizia della regex (rimozione dei delimitatori JS/PHP se presenti)
    clean_pattern = regex_pattern.strip('/')
    if clean_pattern.endswith('/ui'):
        clean_pattern = clean_pattern[:-3]

    # 2. Testiamo l'estrazione del numero di ruolo dall'oggetto o dal corpo
    testo_da_scansionare = f"{oggetto} {corpo}"
    match = re.search(clean_pattern, testo_da_scansionare, re.IGNORECASE)
    
    if match:
        num_ruolo = match.group(1)
        anno = match.group(2)
        print(f"🎯 Trovato riferimento! Ruolo: RG {num_ruolo}/{anno}")
        
        # 3. Verifichiamo se il fascicolo esiste già su Supabase
        fascicolo_res = supabase.table("fascicoli").select("id, numero_ruolo, anno").eq("numero_ruolo", num_ruolo).eq("anno", anno).execute()
        
        if fascicolo_res.data:
            fascicolo = fascicolo_res.data[0]
            confidenza = 100
            stato_elab = "associato_automatico" if confidenza >= soglia_auto else "associato_manuale"
            
            print(f"✅ Fascicolo esistente trovato nel DB (ID: {fascicolo['id']}). Confidenza: {confidenza}%")
            
            # Registriamo la comunicazione PEC collegata al fascicolo
            salva_pec_su_db(
                fascicolo_id=fascicolo['id'],
                mittente=mittente,
                oggetto=oggetto,
                corpo=corpo,
                stato=stato_elab,
                confidenza=confidenza,
                motivo=f"Match esatto RG {num_ruolo}/{anno} da oggetto/testo"
            )
        else:
            print(f"⚠️ Rilevato RG {num_ruolo}/{anno} ma il fascicolo non esiste ancora nel DB. Va in coda di creazione.")
            # Qui l'agente può demandare la creazione della cartella/fascicolo
    else:
        print("❌ Nessun numero di ruolo intercettato. La PEC finisce nella fascia del 'rumore' / gestione manuale.")
        salva_pec_su_db(
            fascicolo_id=None,
            mittente=mittente,
            oggetto=oggetto,
            corpo=corpo,
            stato="scartato_rumore",
            confidenza=0,
            motivo="Nessun pattern RG riconosciuto"
        )

def salva_pec_su_db(fascicolo_id, mittente, oggetto, corpo, stato, confidenza, motivo):
    """Salva l'esito della comunicazione PEC nella tabella comunicazioni_pec di Supabase"""
    data = {
        "fascicolo_id": fascicolo_id,
        "mittente": mittente,
        "oggetto": oggetto,
        "corpo_testo": corpo,
        "stato_elaborazione": stato,
        "confidenza_match": confidenza,
        "motivo_match": motivo
    }
    supabase.table("comunicazioni_pec").insert(data).execute()
    print("💾 Dati PEC salvati correttamente nel database cloud.")

if __name__ == "__main__":
    # Test di simulazione con una PEC in arrivo
    analizza_e_smista_pec(
        mittente="cancelleria.giustizia@pec.it",
        oggetto="Deposito Ordinanza - R.G. 1234/2026 - INPS c/ Rossi",
        corpo="Si comunica il deposito della sentenza relativa al procedimento RG 1234 del 2026."
    )
