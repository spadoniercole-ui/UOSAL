"""
supabase_writer.py
==================
Scrive su Supabase (schema registro) quello che l'ingestione ha ricavato
dalle PEC. Modulo separato da pec_intake.py per due motivi:
  - si testa da solo:  python supabase_writer.py
  - si puo' disattivare senza toccare l'ingestione (basta togliere
    SUPABASE_KEY dall'ambiente e la scrittura si spegne da sola)

Non usa httpx ne' requests: solo urllib, che e' nella libreria standard.
Se manca la chiave o la rete, la scrittura si disattiva con un avviso e
l'ingestione continua a lavorare sui file come prima.
"""

import json
import os
import re
import sys
import uuid
from datetime import datetime, timezone
from urllib import request as _ur, error as _ue
from urllib.parse import quote as _url_quote

# ---------------------------------------------------------------- configurazione
SUPABASE_URL = os.environ.get("SUPABASE_URL", "https://cjfkqhwoxbjeznarvazm.supabase.co")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "sb_publishable_XJlz6PVJeba0EX9BZFDUbA_qeHIXu7f")
SCHEMA       = os.environ.get("SUPABASE_SCHEMA", "registro")

DISABILITATO = not SUPABASE_KEY

# Windows usa cp1252 di default: alcuni caratteri (frecce, accenti rari)
# fanno esplodere il print. Qui si forza UTF-8 sull'output, con fallback
# silenzioso se non e' possibile.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def log(m):  print(f"[{datetime.now():%H:%M:%S}] db: {m}")
def err(m):  print(f"[{datetime.now():%H:%M:%S}] db ERRORE: {m}", file=sys.stderr)


def _headers():
    return {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "resolution=merge-duplicates,return=minimal",
        "Accept-Profile":  SCHEMA,
        "Content-Profile": SCHEMA,
    }


def _chiama(metodo, percorso, corpo=None):
    """Una POST/PATCH/GET verso PostgREST. Ritorna (status, testo)."""
    if DISABILITATO:
        return 0, "scrittura disattivata: manca SUPABASE_KEY"
    url = f"{SUPABASE_URL}/rest/v1/{percorso}"
    dati = None
    if corpo is not None:
        dati = json.dumps(corpo, default=str, ensure_ascii=False).encode("utf-8")
    req = _ur.Request(url, data=dati, headers=_headers(), method=metodo)
    try:
        with _ur.urlopen(req, timeout=30) as r:
            return r.status, r.read().decode("utf-8", "ignore")
    except _ue.HTTPError as e:
        return e.code, e.read().decode("utf-8", "ignore")
    except Exception as e:
        return -1, f"{type(e).__name__}: {e}"


def _upsert(tabella, riga):
    st, tx = _chiama("POST", tabella, riga)
    if st in (200, 201, 204):
        return True
    err(f"{tabella}: HTTP {st} \u2014 {tx[:200]}")
    return False


def _select(tabella, filtro, campi="id"):
    st, tx = _chiama("GET", f"{tabella}?select={campi}&{filtro}")
    if st != 200:
        err(f"lettura {tabella}: HTTP {st}")
        return []
    try:
        return json.loads(tx)
    except Exception:
        return []


def _q(v):
    """URL-encoding per i valori nelle query string di PostgREST."""
    return _url_quote(str(v), safe="")


def _adesso():
    """Timestamp UTC in ISO 8601, come lo vuole Postgres per timestamptz."""
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


# ---------------------------------------------------------------- lookup
def _id_fascicolo(sisco=None, rg=None):
    """Trova l'id di un fascicolo per numero di pratica (preferito) o ruolo.

    Il SISCO viene scritto in due formati: '735/26/CO/1' (archivio storico)
    e '0735/26/CO/1' (pec_intake con zero davanti). Vanno provati entrambi,
    altrimenti la patch di pec_intake crea fascicoli duplicati.
    """
    if not (sisco or rg):
        return None
    if sisco:
        # 1. come arriva
        r = _select("fascicoli", f"sisco=eq.{_q(sisco)}")
        if r:
            return r[0]["id"]
        # 2. senza zeri davanti
        sisco_norm = re.sub(r"^0+", "", str(sisco))
        if sisco_norm and sisco_norm != sisco:
            r = _select("fascicoli", f"sisco=eq.{_q(sisco_norm)}")
            if r:
                return r[0]["id"]
        # 3. con zeri a 4 cifre davanti
        m = re.match(r"^(\d+)(.*)$", str(sisco))
        if m:
            sisco_zero = m.group(1).zfill(4) + m.group(2)
            if sisco_zero != sisco:
                r = _select("fascicoli", f"sisco=eq.{_q(sisco_zero)}")
                if r:
                    return r[0]["id"]
    if rg:
        r = _select("fascicoli", f"rg=eq.{_q(rg)}")
        if r:
            return r[0]["id"]
    return None


def _id_giudice(nome):
    """Trova l'id di un giudice dal nome (match tollerante sul cognome)."""
    if not nome:
        return None
    cognome = nome.strip().upper().split()[0]
    r = _select("giudici", f"nome=ilike.*{_q(cognome)}*")
    return r[0]["id"] if r else None


# ---------------------------------------------------------------- scrittura fascicolo
def _nome_fascicolo(scheda, trib, esistente):
    """Il nome del fascicolo, senza sovrascrivere un nome buono con
    "DA COMPLETARE".

    Se la PEC non porta il nome E il fascicolo esiste gia', restituire None
    significa "non toccare": il filtro `if v is not None` in fondo a
    scrivi_fascicolo scarta la chiave, e l'UPSERT non la aggiorna.
    """
    nome = (scheda.get("ricorrente") or trib.get("parte") or "").strip()
    if nome:
        return nome.upper()
    if esistente:
        return None        # il fascicolo ha gia' un nome: non lo tocchiamo
    return "DA COMPLETARE"


def scrivi_fascicolo(scheda, rif, ufficio, trib, quando):
    """UPSERT del fascicolo. Ritorna l'id (nuovo o esistente).

    scheda  la scheda_pec.json prodotta dall'ingestione
    rif     il dizionario dei riferimenti (rg, sisco, sentenza, ...)
    ufficio {"sede": ..., "grado": ...}
    trib    il modulo di cancelleria, se c'e'
    quando  data di ricezione della PEC
    """
    rif = rif or {}
    trib = trib or {}

    # se esiste gia', non lo ricreo: lo aggiorno solo dove serve
    esistente = _id_fascicolo(rif.get("sisco"), rif.get("rg"))
    nuovo_id  = esistente or ("N" + uuid.uuid4().hex[:10].upper())

    riga = {
        "id": nuovo_id,
        "ricorrente": _nome_fascicolo(scheda, trib, esistente),
        # se il fascicolo esiste gia', non riscrivere il sisco: potrebbe
        # avere un formato diverso (storico senza zero) e non va cambiato
        "sisco": None if esistente else rif.get("sisco"),
        "rg": rif.get("rg"),
        "sede": ufficio.get("sede") or trib.get("sede"),
        "grado": ufficio.get("grado"),
        "materia_testo": trib.get("materia"),
        "giudice_testo": trib.get("giudice") or scheda.get("giudice"),
        "data_pec": str(quando)[:10] if quando else None,
        "stato": "attivo",
        "da_pec": scheda.get("chiave"),
        "cartella_pec": scheda.get("cartella"),
        "note": f"creato dall'ingestione PEC del {str(quando)[:10]}"
                if not esistente else None,
    }
    # il giudice_id e' un FK: si risolve solo se lo troviamo in anagrafica
    gid = _id_giudice(riga.get("giudice_testo"))
    if gid:
        riga["giudice_id"] = gid

    riga = {k: v for k, v in riga.items() if v is not None}

    if _upsert("fascicoli", riga):
        log(("aggiornato" if esistente else "creato") +
            f" fascicolo {nuovo_id} \u2014 "
            f"{(riga.get('ricorrente') or '(nome invariato)')[:40]}") 
        return nuovo_id
    return esistente


# ---------------------------------------------------------------- comunicazione
def _riga_comunicazione(chiave, scheda, rif, fascicolo_id):
    """Costruisce la riga per registro.comunicazioni rispettando i NOT NULL:
    ricevuta, riferimenti, novita, creata_il devono esserci SEMPRE."""
    # ricevuta: se non c'e', meglio un timestamp odierno che un errore
    ricevuta = scheda.get("ricevuta")
    if not ricevuta:
        ricevuta = _adesso()
    # normalizza a ISO con timezone se e' una stringa semplice
    if isinstance(ricevuta, str) and "T" not in ricevuta:
        ricevuta = ricevuta.replace(" ", "T") + "+00:00"

    return {
        "chiave":           chiave,
        "ricevuta":         ricevuta,
        "fascicolo_id":     fascicolo_id,
        "tipo":             scheda.get("tipo_pec") or scheda.get("tipo") or "da_classificare",
        "fase_ctu":         scheda.get("fase_ctu"),
        "oggetto":          scheda.get("oggetto"),
        "oggetto_reale":    scheda.get("oggetto_reale") or scheda.get("oggetto"),
        "oggetto_certificato": scheda.get("oggetto_certificato"),
        "testo":            (scheda.get("testo") or "")[:200000] or None,
        "mittente":         scheda.get("mittente"),
        "mittente_reale":   scheda.get("mittente_reale"),
        "mittente_pec":     scheda.get("mittente_pec"),
        "riferimenti":      rif or {},
        "novita":           scheda.get("novita") or [],
        "notifica_valida":  scheda.get("notifica_valida"),
        "casella_notifica": scheda.get("casella_notifica"),
        "termine_caricamento": scheda.get("termine_caricamento"),
        "accantonata_in":   scheda.get("accantonata_in"),
        "da_recuperare":    bool(scheda.get("da_recuperare")),
        "creata_il":        _adesso(),
    }


def scrivi_comunicazione(scheda, rif=None, ufficio=None, trib=None, quando=None,
                         fascicolo_id=None):
    """UPSERT della PEC in registro.comunicazioni."""
    chiave = scheda.get("chiave")
    if not chiave:
        return False

    # PATCH COPIA: se la riga esiste, aggiorna i campi mancanti
    # (evita che le righe create dal ramo copia restino per sempre vuote)
    esistente = _select("comunicazioni", f"chiave=eq.{_q(chiave)}",
                        campi="chiave,testo,documenti,novita,ricorrente")
    if esistente:
        ex = esistente[0]
        aggiornamenti = {}
        campi_da_verificare = {
            "testo": scheda.get("testo"),
            "documenti": scheda.get("documenti") or [],
            "novita": scheda.get("novita") or [],
            "ricorrente": scheda.get("ricorrente"),
            "sede": scheda.get("sede"),
            "grado": scheda.get("grado"),
            "giudice": scheda.get("giudice"),
            "data_udienza": scheda.get("data_udienza"),
        }
        for k, v in campi_da_verificare.items():
            if v and (not ex.get(k) or ex.get(k) in (None, "", [])):
                aggiornamenti[k] = v
        if aggiornamenti:
            aggiornamenti["chiave"] = chiave
            if _upsert("comunicazioni", aggiornamenti):
                log(f"comunicazione {chiave[:36]} \u2192 aggiornata ({', '.join(aggiornamenti.keys())})")
        return True

    rif  = rif  or scheda.get("riferimenti") or {}
    trib = trib or scheda.get("comunicazione_tribunale") or {}

    if fascicolo_id is None:
        fascicolo_id = _id_fascicolo(rif.get("sisco"), rif.get("rg"))

    riga = _riga_comunicazione(chiave, scheda, rif, fascicolo_id)
    # PATCH COPIA: aggiungi esplicitamente i campi che _riga_comunicazione
    # potrebbe non includere
    if scheda.get("testo") and not riga.get("testo"):
        riga["testo"] = scheda["testo"]
    if scheda.get("documenti") and not riga.get("documenti"):
        riga["documenti"] = scheda["documenti"]
    if scheda.get("novita") and not riga.get("novita"):
        riga["novita"] = scheda["novita"]
    if scheda.get("ricorrente") and not riga.get("ricorrente"):
        riga["ricorrente"] = scheda["ricorrente"]
    if scheda.get("sede") and not riga.get("sede"):
        riga["sede"] = scheda["sede"]
    if scheda.get("grado") and not riga.get("grado"):
        riga["grado"] = scheda["grado"]
    if scheda.get("giudice") and not riga.get("giudice"):
        riga["giudice"] = scheda["giudice"]
    if scheda.get("data_udienza") and not riga.get("data_udienza"):
        riga["data_udienza"] = scheda["data_udienza"]
    riga = {k: v for k, v in riga.items() if v is not None}

    if _upsert("comunicazioni", riga):
        log(f"comunicazione {chiave[:36]} \u2192 fascicolo {fascicolo_id or '—'}")
        return True
    return False

# ---------------------------------------------------------------- udienza
def scrivi_udienza(fascicolo_id, data, quando=None, giudice_id=None, chiave_pec=None):
    """Inserisce una udienza se non ce n'e' gia' una identica."""
    if not (fascicolo_id and data):
        return False
    d = str(data)[:10]
    if _select("udienze", f"fascicolo_id=eq.{fascicolo_id}&data=eq.{d}"):
        return True
    riga = {
        "id": "U" + uuid.uuid4().hex[:10].upper(),
        "fascicolo_id": fascicolo_id,
        "data": d,
        "giudice_id": giudice_id,
        "da_comunicazione": chiave_pec,
        "da_riportare_sisco": True,
        "nota": f"letta dall'ingestione PEC del {str(quando)[:10]}" if quando else None,
    }
    riga = {k: v for k, v in riga.items() if v is not None}
    if _upsert("udienze", riga):
        log(f"udienza {d} aggiunta al fascicolo {fascicolo_id}")
        return True
    return False

# ---------------------------------------------------------------- sentenza dai fatti
_TIPI_SENTENZA = {
    "accolto":            "FAVOREVOLE",
    "parziale":           "PARZIALMENTE FAVOREVOLE",
    "rigettato":          "SFAVOREVOLE",
    "inammissibile":      "SFAVOREVOLE",
    "improcedibile":      "SFAVOREVOLE",
    "estinto":            "ESTINZIONE",
    "cessata_materia":    "CESSATA MATERIA",
    "cancellata":         "ESTINZIONE",
    "conciliato":         "OMOLOGA",
    "incompetenza":       "ORDINANZA",
}

def scrivi_sentenza_dai_fatti(fascicolo_id, novita, ricorrente=None, sede=None,
                              chiave_pec=None):
    """Se fra i fatti c'e' una sentenza con numero, la crea in registro.sentenze."""
    if not (fascicolo_id and novita):
        return None
    num = next((f.get("valore") for f in novita
                if f.get("tipo") == "sentenza" and f.get("valore")), None)
    if not num:
        return None
    # gia' presente?
    if _select("sentenze", f"fascicolo_id=eq.{fascicolo_id}&rg_sentenza=eq.{_q(num)}"):
        return None
    esito   = next((f.get("valore") for f in novita if f.get("tipo") == "esito"), None)
    data    = next((f.get("data")   for f in novita if f.get("tipo") == "sentenza"), None)
    riga = {
        "id":           "S" + uuid.uuid4().hex[:10].upper(),
        "fascicolo_id": fascicolo_id,
        "ricorrente":   ricorrente,
        "sede":         sede,
        "rg_sentenza":  num,
        "tipologia":    _TIPI_SENTENZA.get(esito),
        "parziale":     esito == "parziale",
        "data_definizione": data,
        "da_comunicazione": chiave_pec,
        "da_riportare_sisco": True,
    }
    riga = {k: v for k, v in riga.items() if v is not None}
    if _upsert("sentenze", riga):
        log(f"sentenza {num} creata sul fascicolo {fascicolo_id}")
        return riga["id"]
    return None


# ---------------------------------------------------------------- orchestrazione
def scrivi_scheda(scheda, tipo_esito=None):
    """Scrive tutto quello che riguarda una scheda_pec:
       - il fascicolo (nuovo o esistente)
       - l'udienza se ne compare una
       - la sentenza se fra i fatti ce n'e' una
       - la comunicazione
    tipo_esito: 'nuovo' se la PEC ha aperto una cartella, 'coda' se e' finita
                in un fascicolo esistente, 'avviso' se non ha documenti.
    """
    if DISABILITATO:
        return False
    rif     = scheda.get("riferimenti") or {}
    trib    = scheda.get("comunicazione_tribunale") or {}
    ufficio = {"sede": scheda.get("sede"), "grado": scheda.get("grado")}
    quando  = scheda.get("ricevuta")

    fasc_id = _id_fascicolo(rif.get("sisco"), rif.get("rg"))

    # Se la PEC non apre un fascicolo nuovo E non ha ne' SISCO ne' RG,
    # non c'e' identita' per crearlo. Senza questo controllo, ogni avviso
    # orfano produce un fascicolo "DA COMPLETARE" (osservati 5 fantasmi
    # il 19/10/2026). La comunicazione va scritta lo stesso, ma con
    # fascicolo_id = None.
    if tipo_esito == "nuovo":
        fasc_id = scrivi_fascicolo(scheda, rif, ufficio, trib, quando)
    elif not fasc_id and (rif.get("sisco") or rif.get("rg")):
        # Avviso/coda con un riferimento ma il fascicolo non c'e' in DB:
        # ha senso crearlo, rappresenta una PEC nuova che merita un
        # fascicolo.
        fasc_id = scrivi_fascicolo(scheda, rif, ufficio, trib, quando)

    # niente "if not fasc_id: return False": la comunicazione va scritta
    # anche quando il fascicolo e' assente (avviso orfano).

    ud = scheda.get("data_udienza")
    if ud:
        scrivi_udienza(fasc_id, ud, quando=quando, chiave_pec=scheda.get("chiave"))

    scrivi_sentenza_dai_fatti(fasc_id, scheda.get("novita"),
                              ricorrente=scheda.get("ricorrente"),
                              sede=ufficio.get("sede") or trib.get("sede"),
                              chiave_pec=scheda.get("chiave"))

    scrivi_comunicazione(scheda, rif, ufficio, trib, quando, fascicolo_id=fasc_id)
    return True


# ---------------------------------------------------------------- PEC orfane (senza fascicolo)
def _scrivi_comunicazione_orfana(scheda, marcatura=None):
    """Scrive in registro.comunicazioni una PEC che NON ha fascicolo:
    un avviso del Portale, una comunicazione senza documenti, o una
    accantonata. Non tocca la tabella fascicoli."""
    chiave = scheda.get("chiave")
    if not chiave:
        return False
    if _select("comunicazioni", f"chiave=eq.{_q(chiave)}", campi="chiave"):
        return True
    rif = scheda.get("riferimenti") or {}
    riga = _riga_comunicazione(chiave, scheda, rif, fascicolo_id=None)
    if marcatura:
        riga["accantonata_in"] = marcatura
    riga = {k: v for k, v in riga.items() if v is not None}
    return _upsert("comunicazioni", riga)


def scrivi_avviso_archiviato(scheda):
    """Un avviso del Portale, senza documenti e senza fascicolo."""
    return _scrivi_comunicazione_orfana(scheda)


def scrivi_accantonato_archiviato(scheda):
    """Posta interna messa da parte: stessa scrittura, con la marcatura."""
    return _scrivi_comunicazione_orfana(scheda, marcatura="interno")

def scrivi_copia_archiviata(scheda):
    """La copia della stessa comunicazione, arrivata a un'altra casella.
    E' una PEC senza fascicolo proprio: il legame con l'originale resta nel
    scheda_pec.json su disco (voce 'copie'), non serve una colonna dedicata.
    Stessa scrittura dell'avviso, ma senza la marcatura 'interno'."""
    return _scrivi_comunicazione_orfana(scheda, marcatura=None)

def scrivi_copia_archiviata(scheda):
    """La copia della stessa comunicazione, arrivata a un'altra casella.
    E' una PEC senza fascicolo proprio: il legame con l'originale resta nel
    scheda_pec.json su disco (voce 'copie'), non serve una colonna dedicata.
    Stessa scrittura dell'avviso, ma senza la marcatura 'interno'."""
    return _scrivi_comunicazione_orfana(scheda, marcatura=None)

# ---------------------------------------------------------------- test da riga di comando
def _test():
    print(f"URL: {SUPABASE_URL}")
    print(f"schema: {SCHEMA}")
    print(f"chiave: {'presente' if SUPABASE_KEY else 'MANCANTE'}")
    st, tx = _chiama("GET", "fascicoli?select=id&limit=1")
    print(f"prova lettura fascicoli: HTTP {st}")
    st2, tx2 = _chiama("GET", "comunicazioni?select=chiave&limit=1")
    print(f"prova lettura comunicazioni: HTTP {st2}")
    if st == 200 and st2 == 200:
        print("connessione OK")
        return 0
    print(f"risposta fascicoli: {tx[:200]}")
    print(f"risposta comunicazioni: {tx2[:200]}")
    return 1


if __name__ == "__main__":
    sys.exit(_test())