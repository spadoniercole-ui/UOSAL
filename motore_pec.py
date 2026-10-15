import re
from supabase import create_client, Client

# Credenziali Supabase e chiave memorizzata
SUPABASE_URL = "https://cjfkqhwoxbjeznarvazm.supabase.co"
SUPABASE_KEY = "sb_publishable_pZ3MqT04XNr6C78tHfCBBg_mVp5wfgS"

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

def carica_parametri():
    """Carica le configurazioni e le regex direttamente dalla tabella parametrica su Supabase"""
    response = supabase.table("parametri_sistema").select("chiave, valore").execute()
    return {row["chiave"]: row["valore"] for row in response.data}

def analizza_e_smista_pec(mittente: str, oggetto: str, corpo: str, username_operatore: str = "e.spadoni"):
    """Motore logico per il parsing, il matching della PEC e l'associazione anagrafica"""
    print(f"\n--- Analisi nuova PEC da: {mittente} ---")
    print(f"Oggetto: {oggetto}")
    
    # 1. Recuperiamo l'ID dell'utente responsabile basato sullo username standard
    user_res = supabase.table("utenti").select("id, sede_id, username").eq("username", username_operatore).execute()
    if not user_res.data:
        print(f"⚠️ Attenzione: Utente '{username_operatore}' non trovato nel database.")
        return
    
    operatore = user_res.data[0]
    sede_id_operatore = operatore["sede_id"]
    utente_id = operatore["id"]
    
    # 2. Carichiamo i parametri dinamici dal database
    parametri = carica_parametri()
    soglia_auto = int(parametri.get("soglia_match_automatico", 85))
    regex_pattern = parametri.get("regex_ruolo_principale", r'/(?:R\.?G\.?|RG|RUOLO)\s*(?:N\.?|NUM\.?\s*)?(\d{1,6})[\/\-](\d{2,4})/ui')
    
    clean_pattern = regex_pattern.strip('/')
    if clean_pattern.endswith('/ui'):
        clean_pattern = clean_pattern[:-3]

    # 3. Testiamo l'estrazione del numero di ruolo
    testo_da_scansionare = f"{oggetto} {corpo}"
    match = re.search(clean_pattern, testo_da_scansionare, re.IGNORECASE)
    
    if match:
        num_ruolo = match.group(1)
        anno = match.group(2)
        print(f"🎯 Trovato riferimento! Ruolo: RG {num_ruolo}/{anno}")
        
        # 4. Verifichiamo se il fascicolo esiste già su Supabase
        fascicolo_res = supabase.table("fascicoli").select("id, numero_ruolo, anno, sede_id").eq("numero_ruolo", num_ruolo).eq("anno", anno).execute()
        
        if fascicolo_res.data:
            fascicolo = fascicolo_res.data[0]
            confidenza = 100
            stato_elab = "associato_automatico" if confidenza >= soglia_auto else "associato_manuale"
            
            print(f"✅ Fascicolo esistente trovato nel DB (ID: {fascicolo['id']}). Confidenza: {confidenza}%")
            
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
            print(f"⚠️ Rilevato RG {num_ruolo}/{anno}: Il fascicolo non esiste. Creazione automatica in corso...")
            
            # Creiamo il nuovo fascicolo agganciandolo alla sede dell'operatore e all'utente responsabile
            nuovo_fascicolo = {
                "numero_ruolo": num_ruolo,
                "anno": int(anno),
                "materia": "Contenzioso Generale",
                "stato": "attivo",
                "sede_id": sede_id_operatore,
                "utente_responsabile_id": utente_id,
                "note": f"Creato automaticamente da PEC ricevuta il {mittente}"
            }
            
            ins_res = supabase.table("fascicoli").insert(nuovo_fascicolo).execute()
            if ins_res.data:
                nuovo_id = ins_res.data[0]["id"]
                print(f"📁 Fascicolo RG {num_ruolo}/{anno} creato con successo (ID: {nuovo_id}) associato all'utente @{username_operatore}!")
                
                # Salviamo la PEC collegandola al nuovo fascicolo appena creato
                salva_pec_su_db(
                    fascicolo_id=nuovo_id,
                    mittente=mittente,
                    oggetto=oggetto,
                    corpo=corpo,
                    stato="associato_ automatico",
                    confidenza=100,
                    motivo="Creazione automatica fascicolo da PEC"
                )
    else:
        print("❌ Nessun numero di ruolo intercettato. La PEC finisce nella fascia del 'rumore'.")
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
    # Test di simulazione con una PEC che cita un fascicolo NUOVO (per testare la creazione automatica)
    analizza_e_smista_pec(
        mittente="cancelleria.tribunale@pec.it",
        oggetto="Atto di Citazione - R.G. 9988/2026 - INPS c/ Controparte",
        corpo="Si notifica l'atto di citazione relativo al procedimento RG 9988 del 2026.",
        username_operatore="e.spadone"
    )
EOF