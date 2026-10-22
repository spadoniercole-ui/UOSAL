"""
INGESTIONE PEC - ufficio legale
================================
Legge la posta di Outlook, riconosce le PEC che interessano, le atomizza
e deposita i documenti in una cartella di transito.

NON usa regole di Outlook: i criteri stanno in config_pec.json, sono
ispezionabili e non possono disattivarsi da soli.

Due modalita':

    python pec_intake.py --ricognizione
        Legge e basta. Non scrive nulla se non un referto.
        Da usare la prima volta per vedere cosa c'e' in casella.

    python pec_intake.py
        Elabora davvero: estrae, deposita, aggiorna l'archivio.

    python pec_intake.py --rileggi
        Ignora l'archivio e rielabora tutto (per ricostruire dopo modifiche).

Richiede: pywin32  (pip install pywin32)
Opzionale: asn1crypto per i .p7m (altrimenti usa Windows)
"""

import argparse
import csv
import hashlib
import json
import os
import re
import shutil
import stat
import time
import subprocess
import sys
import unicodedata
import zipfile
import email
import os
from supabase import create_client, Client

# Lettura delle variabili d'ambiente
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
PEC_USER = os.environ.get("PEC_USER")
PEC_PASSWORD = os.environ.get("PEC_PASSWORD")
PEC_IMAP_SERVER = os.environ.get("PEC_IMAP_SERVER")

# Esempio di inizializzazione
if not SUPABASE_URL or not SUPABASE_KEY:
    raise ValueError("Mancano le credenziali di Supabase!")

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
from base64 import b64encode
from datetime import date, datetime, timedelta
from email import policy
from pathlib import Path

BASE = Path(__file__).resolve().parent

# --- Scrittura sul database (schema registro) ---
# Se supabase_writer.py manca, o la libreria di rete non funziona, la
# scrittura si spegne con un avviso e l'ingestione continua a lavorare sui
# file come prima. Non e' un errore: e' un modo di lavorare piu' lento.
try:
    import supabase_writer as DB
    DB_ATTIVO = not DB.DISABILITATO
    if not DB_ATTIVO:
        print("[db] SUPABASE_KEY mancante: scrittura sul database disattivata")
except ImportError as _e:
    DB = None
    DB_ATTIVO = False
    print(f"[db] supabase_writer non disponibile ({_e}): si lavora solo su file")

TIPI = {}
CASELLE = {}
FASI = {}
VERSIONE = "2026-10-01e sbarramenti-al-posto-giusto (i due controlli sulla finestra stavano prima dei comandi che non leggono posta: '--azzera tutto' si vedeva rifiutare con un messaggio che non c'entrava. Adesso stanno dove il giro comincia davvero, e i dieci comandi di servizio - azzera, promuovi, converti, pulisci, lanci, annulla, cerca, raccogli, riclassifica, leggi-cartelle - escono prima. Contiene finestra-dichiarata, ricognizione-vede-il-modulo, descrizione-cancelleria, ruolo-possibile-dal-registro e outlook-aspetta-la-rete - passo con registro 1.112)"
CFG = {}
PER_CARTELLA = []      # quante PEC per cartella, anche zero
UDIENZE_TROVATE = []   # udienze lette dalle PEC per fascicoli che non le hanno
ACCANTONA = True       # sposta in ALTRI GIUDICI / DA ESAMINARE
ESECUZIONI = True      # mette da parte le esecuzioni: le lavora un altro ufficio

# La sezione si nomina nell'oggetto o nel modulo di cancelleria: "SEZIONE
# ESECUZIONI MOBILIARI", "esecuzioni immobiliari", "pignoramento presso terzi".
# E' una regola per forma - il nome dell'ufficio giudiziario - non un elenco
# di parole chiave: se domani cambia il numero di ruolo o l'avvocato, regge.
RE_ESECUZIONI = re.compile(
    r"sezione\s+esecuzion\w*"
    r"|esecuzion\w*\s+(?:mobiliar|immobiliar)\w*"
    r"|pignoramento\s+(?:presso\s+terzi|ex\s+art)"
    r"|R\.?G\.?E\.?\s*n?\.?\s*\d+", re.I)

# AVVISI DI ATTO DA RECUPERARE.
# Certe notifiche non portano niente: dicono che sui server dell'Istituto c'e'
# un atto, e che deve andarlo a prendere una persona autorizzata. Non hanno
# allegati e non possono avere riferimenti, perche' l'informazione sta di la',
# non nel messaggio.
# Non sono rumore e non sono comunicazioni incomplete: sono lavoro. Metterle
# fra le scartate voleva dire nasconderle, e sono 983 in archivio.
RE_AVVISO_ATTO = re.compile(
    r"^\s*notific\w*\s+(?:ai sensi|ex)|"
    r"^\s*notific\w*\s+.{0,20}\b(?:l\.?\s*)?53", re.I)

def e_avviso_da_recuperare(ogg, mittente, allegati, rif):
    """L'atto c'e', ma non qui: va recuperato dal portale.

    Si riconosce per forma, e tutte e quattro le condizioni devono valere:
      - l'oggetto annuncia una notificazione;
      - il mittente e' interno o e' il sistema che avvisa;
      - non ci sono allegati, quindi l'atto non viaggia col messaggio;
      - non c'e' nessun riferimento, quindi non si puo' agganciare a niente.
    Una notificazione vera - con la busta o col ruolo - non la tocca: in
    archivio ce ne sono 557 e restano dove sono."""
    if not RE_AVVISO_ATTO.match(testo(ogg)):
        return False
    if allegati:
        return False
    # "rg_forse" e' un ruolo POSSIBILE - il numero di COMUNICAZIONE, che meta'
    # delle volte e' un ruolo - non un riferimento vero. Contarlo qui faceva
    # sparire tutti gli avvisi da prelevare: hanno l'oggetto "COMUNICAZIONE
    # n/anno/LAV" e quindi un rg_forse sempre.
    if any(v for k, v in (rif or {}).items() if k != "rg_forse"):
        return False
    m = testo(mittente).lower()
    return bool(re.search(r"@(inps\.it|itinps\.onmicrosoft\.com"
                          r"|portalecomunicazione\.inps\.it)", m))

def e_esecuzione(*testi):
    """Vero se la comunicazione viene dalla sezione esecuzioni."""
    for t in testi:
        if t and RE_ESECUZIONI.search(testo(t)):
            return True
    return False
SIMULAZIONE = False
SCARTA_INOLTRI = True
IDENTIFICATIVI = {}
ERRORI = []
SCARTA_INTERNE = True
COPIE = {}
# Le firme nate IN QUESTO lancio: serve a dire, nel referto, se una copia
# duplica una PEC vista adesso o una comunicazione gia' in archivio da prima.
FIRME_LANCIO = set()
AVVISI = []
ACCANTONATI = []   # avvisi di provenienza interna: messi da parte, non buttati
LAVORATE = []
IN_CODA = []
SOLO_CON_DOCUMENTI = True
CONFIG = BASE / "config_pec.json"
PERCORSI = BASE / "percorsi.json"   # solo di questa macchina: non viene mai sovrascritto
ARCHIVIO = BASE / "archivio_pec.json"

CONFIG_DEFAULT = {
    "_nota": "Criteri di selezione delle PEC. Modificabili senza toccare il codice.",
    "cartelle": [],
    "_nota_cartelle": "Vuoto = tutte le cartelle della casella. Altrimenti: ['Posta in arrivo', 'LEGALE/00 DA LAVORARE']",
    "giorni_indietro": 90,
    "destinazione": "IN ARRIVO",
    "criteri": {
        "mittenti": ["pec.it", "legalmail.it", "postecert.it", "pec.giustizia.it",
                     "sicurezzapostale.it", "arubapec.it", "pec.aruba.it"],
        "oggetto": ["ricorso", "r.g.", "rg n", "tribunale", "sezione lavoro",
                    "giudice del lavoro", "udienza", "rinvio", "verbale",
                    "sentenza", "decreto", "notifica", "comparizione",
                    "opposizione", "appello", "c.t.u.", "ctu"],
        "corpo": ["notifico", "si trasmette", "in allegato", "ricorso",
                  "decreto di fissazione", "verbale di udienza"],
        "escludi_oggetto": ["newsletter", "fattura elettronica", "posta ordinaria",
                            "spam", "pubblicit"]
    },
    "_nota_criteri": "Basta che scatti UNO fra mittenti/oggetto/corpo. Gli escludi_oggetto hanno la precedenza."
}

# --------------------------------------------------------------------------
def log(m):  print(f"[{datetime.now():%H:%M:%S}] {m}")
ULTIMO_ERRORE = ""

def err(m):
    # L'ultimo motivo detto a schermo serve al diario: con l'esecuzione
    # automatica la console non la legge nessuno, e "si e' fermato da solo"
    # non e' una diagnosi. Ogni return 1 di main() e' preceduto da una err().
    global ULTIMO_ERRORE
    ULTIMO_ERRORE = str(m)
    print(f"[{datetime.now():%H:%M:%S}] ERRORE  {m}")

def testo(v):
    return "" if v is None else str(v)

def pulisci_nome(v, max_len=110):
    """Toglie i caratteri vietati nei nomi file, apostrofi compresi:
    su Windows l'apostrofo non e' vietato, ma spezza i comandi PowerShell."""
    v = unicodedata.normalize("NFKC", testo(v)).strip()
    v = re.sub(r"[\\/:*?\"<>|\r\n\t]+", "_", v)
    v = v.replace("'", "_").replace("\u2019", "_")
    v = re.sub(r"\s+", " ", v).strip(" .")
    return v[:max_len] or "senza_nome"

def leggi_json(p, dflt):
    p = Path(p)
    if not p.exists():
        return dflt
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception as e:
        err(f"lettura {p.name}: {e}")
        return dflt

def scrivi_json(p, dati):
    """Windows rifiuta i percorsi oltre 260 caratteri, e i nomi di cartella
    costruiti dall'oggetto della PEC ci arrivano. Il prefisso \\\\?\\ toglie il
    limite; se anche cosi' non si scrive, l'errore dice QUALE percorso, che
    prima non si leggeva perche' veniva troncato."""
    p = Path(p)
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps(dati, ensure_ascii=False, indent=2), encoding="utf-8")
        return
    except (FileNotFoundError, OSError):
        pass
    lungo = Path("\\\\?\\" + str(p.resolve())) if os.name == "nt" else p
    try:
        lungo.parent.mkdir(parents=True, exist_ok=True)
        lungo.write_text(json.dumps(dati, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        err(f"non scritto ({len(str(p))} caratteri di percorso): {p}  -> {e}")

PERCORSI_DEFAULT = {
    "_nota": ("Percorsi di QUESTA installazione. Tenuti fuori da config_pec.json "
              "perche' quello contiene i criteri e viene aggiornato: cosi' un "
              "aggiornamento dei criteri non cancella i tuoi percorsi."),
    "destinazione": "",
    "_nota_destinazione": "cartella di transito delle PEC, es. .../UOSAL/IN ARRIVO",
    "archivio_fascicoli": "",
    "_nota_archivio_fascicoli": "cartella dei fascicoli fisici, es. .../RICORSI ES",
    "indice_cartelle": "",
    "_nota_indice_cartelle": "indice_cartelle.json esportato dal registro",
    "base_json": "",
    "_nota_base_json": "base.json del registro fascicoli",
}

def cerca_percorsi():
    """Se i percorsi non sono indicati, provo a trovarli da solo risalendo le
    cartelle: base.json e RICORSI ES hanno nomi riconoscibili."""
    trovati = {}
    radice = BASE
    for _ in range(4):
        if not radice.exists():
            break
        for d in list(radice.iterdir()) + [radice]:
            try:
                if not d.is_dir():
                    continue
            except Exception:
                continue
            if not trovati.get("base_json") and (d / "base.json").exists():
                trovati["base_json"] = str(d / "base.json")
                if (d / "indice_cartelle.json").exists():
                    trovati["indice_cartelle"] = str(d / "indice_cartelle.json")
                if not trovati.get("destinazione") and (d / "IN ARRIVO").exists():
                    trovati["destinazione"] = str(d / "IN ARRIVO")
            if not trovati.get("archivio_fascicoli") and \
               re.search(r"ricorsi", d.name, re.I) and d != radice:
                trovati["archivio_fascicoli"] = str(d)
        if trovati.get("base_json") and trovati.get("archivio_fascicoli"):
            break
        radice = radice.parent
    return trovati

# --------------------------------------------------------------------------
# BUSTE FIRMATE .p7m
# --------------------------------------------------------------------------
def _der_contenuto(dati):
    """Estrae il contenuto originale da una busta PKCS#7 leggendo il DER.
    Nessuna dipendenza esterna, nessun processo PowerShell: quindi nessun
    problema con apostrofi o percorsi strani."""
    def leggi_tlv(b, i):
        if i >= len(b):
            return None
        tag = b[i]; i += 1
        if i >= len(b):
            return None
        n = b[i]; i += 1
        if n & 0x80:
            k = n & 0x7F
            if k == 0:                      # lunghezza indefinita
                return tag, i, None
            n = int.from_bytes(b[i:i+k], "big"); i += k
        return tag, i, n

    def cerca_econtent(b, i, fine, prof=0):
        while i < fine and prof < 12:
            t = leggi_tlv(b, i)
            if not t:
                return None
            tag, inizio, lung = t
            if lung is None:
                lung = fine - inizio
            fine_el = min(inizio + lung, fine)
            costruito = bool(tag & 0x20)
            # OID di id-data (1.2.840.113549.1.7.1): il contenuto segue
            if tag == 0x06 and b[inizio:fine_el] == bytes.fromhex("2a864886f70d010701"):
                t2 = leggi_tlv(b, fine_el)
                if t2:
                    tag2, ini2, lun2 = t2
                    if tag2 == 0xA0:                       # [0] EXPLICIT
                        t3 = leggi_tlv(b, ini2)
                        if t3 and t3[0] in (0x04, 0x24):   # OCTET STRING
                            tag3, ini3, lun3 = t3
                            if tag3 == 0x04:
                                return b[ini3:ini3 + (lun3 or 0)]
                            pezzi, j = b"", ini3
                            limite = ini3 + (lun3 or (fine - ini3))
                            while j < limite:
                                t4 = leggi_tlv(b, j)
                                if not t4:
                                    break
                                tg, ii, ll = t4
                                if tg == 0x04:
                                    pezzi += b[ii:ii + (ll or 0)]
                                j = ii + (ll or 0)
                            if pezzi:
                                return pezzi
            if costruito:
                r = cerca_econtent(b, inizio, fine_el, prof + 1)
                if r:
                    return r
            i = fine_el
        return None

    return cerca_econtent(dati, 0, len(dati))

def apri_p7m(dati):
    """Restituisce il contenuto della busta, o None. Tre tentativi in ordine
    di robustezza: asn1crypto, lettura DER interna, Windows."""
    try:
        from asn1crypto import cms
        ci = cms.ContentInfo.load(dati)
        c = ci["content"]["encap_content_info"]["content"].native
        if c:
            return c if isinstance(c, bytes) else bytes(c)
    except Exception:
        pass
    try:
        c = _der_contenuto(dati)
        if c:
            return c
    except Exception:
        pass
    try:
        # comando codificato in base64: immune ad apostrofi e caratteri speciali
        tmp_in = BASE / "_tmp_busta.p7m"
        tmp_out = BASE / "_tmp_busta.out"
        tmp_in.write_bytes(dati)
        ps = (f"Add-Type -AssemblyName System.Security;"
              f"$c=New-Object System.Security.Cryptography.Pkcs.SignedCms;"
              f"$c.Decode([IO.File]::ReadAllBytes([Environment]::GetEnvironmentVariable('PIN')));"
              f"[IO.File]::WriteAllBytes([Environment]::GetEnvironmentVariable('POUT'),$c.ContentInfo.Content)")
        import os
        env = dict(os.environ, PIN=str(tmp_in), POUT=str(tmp_out))
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                        "-EncodedCommand", b64encode(ps.encode("utf-16-le")).decode()],
                       capture_output=True, env=env, timeout=60)
        if tmp_out.exists() and tmp_out.stat().st_size:
            c = tmp_out.read_bytes()
            tmp_in.unlink(missing_ok=True); tmp_out.unlink(missing_ok=True)
            return c
        tmp_in.unlink(missing_ok=True)
    except Exception as e:
        err(f"decodifica p7m: {e}")
    return None

def estensione_vera(dati, nome_originale=""):
    """Il .p7m non contiene sempre un PDF: decido dai byte, non dal nome."""
    firme = [(b"%PDF", ".pdf"), (b"PK\x03\x04", ".zip"), (b"\xd0\xcf\x11\xe0", ".doc"),
             (b"\xff\xd8\xff", ".jpg"), (b"\x89PNG", ".png"), (b"II*\x00", ".tif"),
             (b"MM\x00*", ".tif"), (b"{\\rtf", ".rtf"), (b"7z\xbc\xaf", ".7z")]
    for f, e in firme:
        if dati[:len(f)] == f:
            return e
    t = dati[:400].lstrip()
    if t[:5].lower() == b"<?xml":
        return ".xml"
    if b"Content-Type:" in dati[:600] and b"From:" in dati[:600]:
        return ".eml"
    suff = Path(nome_originale).suffixes
    return suff[-2] if len(suff) >= 2 else ".bin"

# --------------------------------------------------------------------------
# ATOMIZZAZIONE
# --------------------------------------------------------------------------
SALTA = {"smime.p7s", "daticert.xml", "postacert.eml"}

def deposita(dati, nome, dest, visti, esito, provenienza=""):
    if dest is None:                     # lettura senza scrittura
        return
    """Scrive un file. Scarta solo i doppioni veri (stesso contenuto E stesso nome):
    un contenuto uguale con nome diverso viene conservato, perche' il nome
    e' cio' che permette di riconoscere il tipo di documento."""
    nome = pulisci_nome(nome)
    impronta = hashlib.sha256(dati).hexdigest()
    chiave_nome = re.sub(r"[^a-z0-9.]", "", nome.lower())
    if (impronta, chiave_nome) in visti:
        esito["doppioni"] += 1
        return None
    gemelli = [n for (h, n) in visti if h == impronta]
    visti.add((impronta, chiave_nome))
    p = dest / nome
    n = 1
    while p.exists():
        p = dest / f"{Path(nome).stem} ({n}){Path(nome).suffix}"
        n += 1
    p.write_bytes(dati)
    voce = {"nome": p.name, "byte": len(dati), "da": provenienza}
    if gemelli:
        voce["contenuto_uguale_a"] = gemelli[0]
    esito["file"].append(voce)
    return p

def scarta_involucri(dati, nome, dest, visti, esito, prof=0, provenienza=""):
    """Toglie un involucro alla volta finche' non trova un file utilizzabile.
    Un PDF gia' leggibile fa zero giri."""
    if prof > 6:
        esito["note"].append(f"annidamento troppo profondo: {nome}")
        return
    basso = nome.lower()

    if basso.endswith(".xml") and Path(basso).name != "daticert.xml":
        # XML della busta telematica: lo leggo E lo conservo come documento
        try:
            d = leggi_xml_atto(dati)
            if d:
                esito.setdefault("xml_atto", {}).update(d)
        except Exception:
            pass
        deposita(dati, nome, dest, visti, esito, provenienza)
        return

    if Path(basso).name == "daticert.xml":
        # Attestazione del gestore. Ce ne puo' essere piu' d'una: ogni inoltro
        # ne genera una nuova, dove il destinatario e' chi ha ricevuto l'inoltro.
        # Le raccolgo tutte, la scelta di quale vale si fa dopo.
        d = leggi_daticert(dati)
        if d:
            d["profondita"] = prof
            esito.setdefault("daticert_tutti", []).append(d)
            esito.setdefault("daticert", {}).update(d)
        else:
            esito["note"].append("daticert.xml presente ma illeggibile")
        return

    if basso.endswith(".p7m"):
        # l'originale firmato si conserva: e' il documento che ha valore legale
        deposita(dati, nome, dest, visti, esito, provenienza)
        dentro = apri_p7m(dati)
        if dentro:
            est = estensione_vera(dentro, nome)
            nuovo = Path(nome).with_suffix("").name
            if not nuovo.lower().endswith(est):
                nuovo += est
            scarta_involucri(dentro, nuovo, dest, visti, esito, prof + 1, nome)
        else:
            esito["note"].append(f"busta non apribile: {nome}")
        return

    if basso.endswith(".zip"):
        deposita(dati, nome, dest, visti, esito, provenienza)
        tmp = dest / "_zip_tmp"
        try:
            tmp.mkdir(exist_ok=True)
            z = tmp / "a.zip"; z.write_bytes(dati)
            with zipfile.ZipFile(z) as zf:
                for info in zf.infolist():
                    if info.is_dir():
                        continue
                    scarta_involucri(zf.read(info), Path(info.filename).name,
                                     dest, visti, esito, prof + 1, nome)
        except Exception as e:
            esito["note"].append(f"zip non apribile ({nome}): {e}")
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
        return

    if basso.endswith(".7z"):
        deposita(dati, nome, dest, visti, esito, provenienza)
        esito["note"].append(f"archivio .7z non apribile senza 7-Zip: {nome}")
        return

    if basso.endswith(".eml"):
        esplodi_eml(dati, dest, visti, esito, prof + 1, nome)
        return

    deposita(dati, nome, dest, visti, esito, provenienza)

def leggi_daticert(dati):
    """daticert.xml e' l'attestazione del gestore: dice mittente, destinatari
    certificati e oggetto originali. E' l'unico posto dove sta scritto a quale
    casella l'atto e' stato davvero notificato: nel corpo della busta compare
    chi ha inoltrato, che e' un'altra cosa."""
    try:
        testo_xml = dati.decode("utf-8", "ignore")
    except Exception:
        return {}
    out = {}
    try:
        import xml.etree.ElementTree as ET
        r = ET.fromstring(testo_xml)
        mitt = r.find(".//mittente")
        if mitt is not None and mitt.text:
            out["mittente"] = mitt.text.strip()
        dest = [d.text.strip() for d in r.findall(".//destinatari") if d.text]
        certificati = [d.text.strip() for d in r.findall(".//destinatari")
                       if d.text and d.get("tipo") == "certificato"]
        if dest:
            out["destinatari"] = dest
        if certificati:
            out["destinatari_certificati"] = certificati
        ogg = r.find(".//oggetto")
        if ogg is not None and ogg.text:
            out["oggetto"] = ogg.text.strip()
        dt = r.find(".//data")
        if dt is not None:
            g = dt.find("giorno"); o = dt.find("ora")
            if g is not None and g.text:
                out["data"] = g.text.strip() + (" " + o.text.strip() if o is not None and o.text else "")
    except Exception:
        # ripiego senza parser: i tag sono semplici
        for tag in ("mittente", "destinatari", "oggetto"):
            v = re.findall(r"<" + tag + r"[^>]*>([^<]+)</" + tag + r">", testo_xml)
            if v:
                out[tag if tag != "destinatari" else "destinatari"] = v if tag == "destinatari" else v[0].strip()
    return out

TAG_UTILI = {
    "rg":        ("numeroruolo", "numruolo", "ruologenerale", "nrg", "numerorg",
                  "numeroruologenerale", "numeroregistro"),
    "giudice":   ("giudice", "magistrato", "nomegiudice", "giudicerelatore"),
    "oggetto":   ("oggetto", "descrizioneoggetto", "codiceoggetto", "tipoatto"),
    "sede":      ("ufficiogiudiziario", "descrizioneufficio", "sede", "tribunale"),
    "sezione":   ("sezione", "sezcoll", "codicesezione"),
    "udienza":   ("dataudienza", "udienza", "dataudienzarinvio"),
}
TAG_PARTE = ("denominazione", "cognome", "ragionesociale", "nominativo",
             "cognomenome", "nomeparte", "denominazioneparte")

def leggi_xml_atto(dati):
    """Nelle buste del processo telematico c'e' sempre un XML con i dati
    strutturati. E' l'ultima fonte certa quando il corpo non dice nulla e i
    PDF sono scansioni: gli schemi cambiano, i nomi dei campi molto meno."""
    try:
        testo_xml = dati.decode("utf-8", "ignore")
    except Exception:
        return {}
    if "<" not in testo_xml[:400]:
        return {}
    out, parti = {}, []
    try:
        import xml.etree.ElementTree as ET
        radice = ET.fromstring(testo_xml)
        for el in radice.iter():
            nome = re.sub(r"^\{[^}]*\}", "", el.tag or "").lower().replace("_", "")
            valore = (el.text or "").strip()
            if not valore:
                continue
            if len(valore) > 120:
                # Nelle buste telematiche il modulo di cancelleria sta dentro
                # <Contenuto>, in CDATA, e supera sempre i 120 caratteri:
                # scartandolo si perdeva l'unico posto dove c'e' il nome della
                # parte. Il corpo Outlook di queste PEC e' vuoto, quindi
                # leggi_comunicazione_tribunale non lo vedeva mai.
                d = leggi_comunicazione_tribunale(valore)
                for kk, vv in (d or {}).items():
                    if vv not in (None, "", False) and kk not in out:
                        out[kk] = vv
                continue
            for chiave, tag in TAG_UTILI.items():
                if nome in tag and chiave not in out:
                    out[chiave] = valore
            if nome in TAG_PARTE:
                parti.append(valore)
            for att, av in (el.attrib or {}).items():
                a = att.lower().replace("_", "")
                for chiave, tag in TAG_UTILI.items():
                    if a in tag and chiave not in out and av.strip():
                        out[chiave] = av.strip()
    except Exception:
        # XML malformato: ripiego sui modelli, che non hanno bisogno di parser
        for chiave, tag in TAG_UTILI.items():
            for tg in tag:
                m = re.search(r"<[^>]*" + tg + r"[^>]*>([^<]{1,120})<", testo_xml, re.I)
                if m and chiave not in out:
                    out[chiave] = m.group(1).strip()
        parti = [m.group(1).strip() for tg in TAG_PARTE
                 for m in re.finditer(r"<[^>]*" + tg + r"[^>]*>([^<]{2,80})<", testo_xml, re.I)]

    if out.get("rg"):
        m = re.search(r"(\d{1,6})\s*[/\-]\s*(\d{2,4})", out["rg"])
        if m:
            anno = m.group(2)
            out["rg"] = f"{int(m.group(1))}/{'20'+anno if len(anno)==2 else anno}"
        elif out["rg"].isdigit():
            out.pop("rg")
    if out.get("udienza"):
        d = trova_udienza("udienza " + out["udienza"]) or trova_udienza(out["udienza"])
        out["udienza"] = d.isoformat() if d else None
    # la parte che interessa non e' l'Istituto
    for p in parti:
        if not _e_inps(p):
            ripulita = _ripulisci_nome_persona(p)
            if ripulita:
                out["parte"] = ripulita
                break
    return {k: v for k, v in out.items() if v}

def esplodi_eml(dati, dest, visti, esito, prof=0, provenienza=""):
    """Apre un messaggio: ne salva il testo e passa gli allegati allo scarto."""
    try:
        msg = email.message_from_bytes(dati, policy=policy.default)
    except Exception as e:
        esito["note"].append(f"messaggio illeggibile: {e}")
        return
    corpi = []
    for parte in msg.walk():
        if testo(parte.get_content_disposition()).lower() == "attachment":
            continue
        if parte.get_content_type() == "text/plain":
            try:
                corpi.append(parte.get_content())
            except Exception:
                p = parte.get_payload(decode=True)
                if p:
                    corpi.append(p.decode(parte.get_content_charset() or "utf-8", "ignore"))
    if corpi:
        esito["testo"] += "\n" + "\n".join(corpi)
    for parte in msg.walk():
        nome = parte.get_filename()
        if not nome:
            continue
        base = Path(nome).name.lower()
        carico = parte.get_payload(decode=True)
        if not carico:
            continue
        if base == "daticert.xml":
            d = leggi_daticert(carico)
            if d:
                esito.setdefault("daticert", {}).update(d)
            continue
        if base in SALTA and prof == 0:
            continue
        scarta_involucri(carico, Path(nome).name, dest, visti, esito, prof, provenienza)

# Le cancellerie usano due modelli. Il primo dice "Ricorr./Resist. principale",
# il secondo "Attore/Conv. principale". Senza le seconde due il campo Giudice
# non si chiudeva mai e si mangiava trecento caratteri, e la parte non usciva
# affatto: e' la causa di gran parte dei nominativi mancanti.
ETICHETTE_TRIB = ["Sez/Coll\\.", "Tipo procedimento", "Numero di Ruolo generale", "Giudice",
                  "Ricorr\\. principale", "Resist\\. principale",
                  "Attore principale", "Conv\\. principale", "Convenuto principale",
                  "Data Evento", "Tipo Evento",
                  "Oggetto", "Descrizione",
                  "Note", "Registrato da", "Si vedano", "Si prega"]
_FINE = "(?=" + "|".join(ETICHETTA + r"\s*:?" for ETICHETTA in ETICHETTE_TRIB) + "|$)"

def _campo_trib(t, etichetta):
    # "Oggetto:" compare sia nell'intestazione della mail sia nel modulo:
    # vale l'ultima, che e' quella del modulo
    trovati = [m for m in re.finditer(etichetta + r"\s*:\s*(.*?)" + _FINE, t, re.I | re.S)]
    if not trovati:
        return None
    v = re.sub(r"\s+", " ", trovati[-1].group(1)).strip(" .-\u2013")
    return v or None

def _parte_riga(t, etichetta):
    """Il nome della parte, in tutte e due le forme in cui il modulo arriva.

    A capo, come lo scrive il Portale:
        Ricorr. principale: ZUPPARDI PAOLO
        Resist. principale: INPS DI SIRACUSA
    oppure schiacciato, come esce dal corpo di una PEC inoltrata, senza a capo
    e senza spazi fra un campo e l'altro:
        ...Giudice: MARESCALCO PAOLORicorr. principale: ZUPPARDI PAOLOResist...

    Tagliare a fine riga funziona solo nella prima forma: nella seconda la
    riga non finisce mai e il nome usciva come "CASSIA CARMELAResist.
    principale: ". Tagliare alla prossima etichetta funziona solo nella
    seconda: nella prima, dopo "Attore principale: NOME" viene "Si da' atto
    che...", che etichetta non e', e il campo correva per trecento caratteri.

    Quindi si prende il tratto piu' corto fra i due: la prima etichetta che
    compare, oppure il primo a capo. Quello che viene prima vince."""
    m = re.search(etichetta + r"\s*:\s*(.{1,160})", testo(t), re.I | re.S)
    if not m:
        return None
    v = m.group(1)
    # dove comincia la prossima etichetta, se c'e'
    tagli = [len(v)]
    e = re.search("|".join(x + r"\s*:" for x in ETICHETTE_TRIB), v, re.I)
    if e:
        tagli.append(e.start())
    nl = re.search(r"[\r\n]", v)
    if nl:
        tagli.append(nl.start())
    v = v[:min(tagli)]
    v = re.sub(r"\s+", " ", v).strip(" .-\u2013")
    return v or None

def _e_inps(v):
    return bool(re.search(r"I\.?\s*N\.?\s*P\.?\s*S", testo(v), re.I)) or \
           "ISTITUTO NAZIONALE" in testo(v).upper()

def leggi_comunicazione_tribunale(t):
    """Le comunicazioni di cancelleria arrivano come modulo a campi fissi nel
    corpo della PEC. Leggerlo da' ruolo, giudice, parti e provvedimento senza
    dover aprire un solo allegato - e funziona anche se i PDF sono scansioni."""
    t = testo(t)
    if not re.search(r"Comunicazione di cancelleria|Notificazione di cancelleria"
                     r"|Numero di Ruolo generale", t, re.I):
        return {}
    inizio = re.search(r"Comunicazione di cancelleria", t, re.I)
    corpo_modulo = t[inizio.end():] if inizio else t
    out = {}
    rg = _campo_trib(corpo_modulo, "Numero di Ruolo generale")
    if rg:
        m = re.search(r"(\d{1,6})\s*/\s*(\d{2,4})", rg)
        if m:
            anno = m.group(2)
            out["rg"] = f"{int(m.group(1))}/{'20'+anno if len(anno)==2 else anno}"
    for chiave, eti in (("giudice", "Giudice"), ("materia", "Tipo procedimento"),
                        ("provvedimento", "Oggetto"), ("descrizione", "Descrizione"),
                        ("sezione", "Sez/Coll\\."), ("registrato_da", "Registrato da")):
        v = _campo_trib(corpo_modulo, eti)
        if v:
            out[chiave] = v

    # Due modelli di modulo: Ricorr./Resist. oppure Attore/Conv.
    ric = _parte_riga(corpo_modulo, "Ricorr\\. principale") \
          or _parte_riga(corpo_modulo, "Attore principale")
    res = _parte_riga(corpo_modulo, "Resist\\. principale") \
          or _parte_riga(corpo_modulo, "Conv\\. principale") \
          or _parte_riga(corpo_modulo, "Convenuto principale")
    # quando ricorre l'Istituto, la parte che ci interessa e' il resistente
    parte = res if _e_inps(ric) else ric
    if parte and not _e_inps(parte):
        out["parte"] = re.sub(r"\s+N\.?\s*Q\.?\s*$", "", parte, flags=re.I).strip()
    # "inps_ricorrente" solo se una parte si e' davvero letta. Scriverlo sempre
    # - anche False - rendeva VERO un risultato vuoto: su un inoltro che porta
    # la sola intestazione del modulo ("Comunicazione di cancelleria" con i
    # campi mangiati dalle virgolette del FWD) la funzione tornava
    # {"inps_ricorrente": False}, che e' un dizionario non vuoto. Il chiamante
    # faceva "if not trib" e quindi NON rileggeva il livello interno, dove il
    # modulo sta per intero: ruolo, giudice, parte e provvedimento si
    # perdevano tutti insieme.
    if ric or res:
        out["inps_ricorrente"] = _e_inps(ric)

    m = re.search(r"Tribunale\s+di\s+([A-Z][A-Za-z\u00c0-\u00ff']+)", t)
    if m:
        out["sede"] = m.group(1).upper()
    # "EMESSO DECRETO DI ACCOGLIMENTO N. 570/2026" -> numero del provvedimento
    m = re.search(r"\bN\.?\s*(\d{1,5})\s*/\s*(\d{4})", out.get("descrizione") or "")
    if m:
        out["numero_provvedimento"] = f"{int(m.group(1))}/{m.group(2)}"
    m = re.search(r"Notificato\s+alla\s+PEC[^\d]{0,20}(\d{1,2}/\d{1,2}/\d{4})", t, re.I)
    if m:
        out["notificato_il"] = m.group(1)
    return out

def modulo_di_cancelleria(*testi):
    """Il modulo, cercato su PIU' LIVELLI e unito campo per campo.

    La comunicazione del tribunale arriva inoltrata: dagli avvocati di sede o
    dai sussidiari. A seconda di come viene inoltrata, la parte descrittiva -
    che dice COSA stanno comunicando, e vale per le sentenze come per i rinvii
    d'udienza e le CTU - sta nel corpo del primo livello o in quello del
    secondo.

    Prima si leggeva un livello alla volta e si teneva il primo che rispondeva
    ("if not trib"): un corpo esterno che portava mezzo modulo bloccava la
    lettura di quello interno, che lo portava intero. Qui si leggono tutti e
    per ogni campo vale il primo valore pieno, nell'ordine in cui i testi
    arrivano: dal piu' interno e attendibile al piu' esterno."""
    out = {}
    for t in testi:
        if not t:
            continue
        d = leggi_comunicazione_tribunale(t)
        for k, v in (d or {}).items():
            if v not in (None, "", False) and k not in out:
                out[k] = v
            # i booleani: valgono anche quando sono falsi, ma solo se il testo
            # che li porta ha letto davvero una parte
            elif k == "inps_ricorrente" and k not in out and d.get("parte"):
                out[k] = v
    return out

def leggi_avviso_portale(t):
    """Gli avvisi del Portale hanno anch'essi etichette fisse. Non danno il
    ruolo, ma danno protocollo, data di ricezione e mittente reale: dati che
    servono al fascicolo e che altrimenti si perdono."""
    t = testo(t)
    if not re.search(r"Comunicazione numero|Data di ricezione INPS|assegnazione PEC", t, re.I):
        return {}
    out = {}
    for chiave, eti in (("protocollo", r"Comunicazione numero"),
                        ("ricevuta_inps", r"Data di ricezione INPS"),
                        ("mittente_reale", r"Mittente\s*Da"),
                        ("oggetto_originale", r"Comunicazione\s*:?\s*Oggetto")):
        m = re.search(eti + r"\s*:\s*(.+?)(?=\n|Comunicazione|Data di|Mittente|Oggetto|$)",
                      t, re.I | re.S)
        if m:
            v = re.sub(r"\s+", " ", m.group(1)).strip(" .:-")
            if v:
                out[chiave] = v
    return out

RE_PRATICA_NOME = re.compile(
    r"Pratica\s*:?\s*\d{1,5}\s*/\s*\d{2}\s*/\s*(?:CO|AO|AI)\s*/?\s*\d?\s*[-\u2013]\s*([^)\n\r]{3,60})",
    re.I)

# Parole che in un nominativo non compaiono mai: preposizioni, congiunzioni e
# i ruoli del processo. Servono a fermare le FRASI pescate dal testo degli
# atti - "AD OGNI MODO", "ALLA GESTIONE SEPARATA SIA DERIVATA", "CONSIGLIERE
# RELATORE E CHE SIA" - che finivano in archivio come ricorrenti.
#
# Attenzione a due cose imparate a spese nostre:
#  - le particelle dei cognomi vanno lasciate stare: LO PRESTI, DI STEFANO,
#    DE LUCA sono nomi veri, e "L'OASI DELLA SPERANZA SOC COOP ONLUS" e'
#    una parte legittima;
#  - NON si usano le desinenze verbali: mezzo cognome siciliano finisce in
#    -ANO o -ATO (CAPUANO, ROMANO, AMATO, TROVATO) e verrebbe scartato.
NOME_APERTURA = {"AD","AL","ALL","ALLA","ALLE","AGLI","DEGLI","DELLE","DALLA","DALLE",
                 "NEL","NELLA","SUL","SULLA","PER","CON","TRA","FRA","COME","DOPO",
                 "SENZA","SOTTO","SOPRA","CHE","NON","OGNI","QUALE","QUALI","IN","E",
                 "O","MA","SE","SI","CI","VI","LORO","QUANTO","MENTRE","PERCHE",
                 "POICHE","AFFINCHE","OVVERO","NONCHE","ALTRESI","PERTANTO","QUINDI"}
NOME_INTERNE = {"AD","AL","ALL","ALLA","ALLE","AGLI","DEGLI","DALLA","DALLE","NEL",
                "NELLA","SUL","SULLA","COME","DOPO","SENZA","SOTTO","SOPRA","CHE",
                "NON","OGNI","QUALE","QUALI","MENTRE","PERCHE","OVVERO","PERTANTO",
                "QUINDI","LORO"}
NOME_RUOLI = {"GIUDICE","CONSIGLIERE","RELATORE","RICORRENTE","RESISTENTE","CONVENUTO",
              "ATTORE","CANCELLIERE","PRESIDENTE","MAGISTRATO","DIFENSORE","PROCURATORE",
              "ANTISTATARI","PROBLEMI","GIUDIZIARIE","FASC"}

# Parole che nominano un DOCUMENTO o un atto, non una persona. Valgono solo
# se stanno IN TESTA al nominativo: "SCATA S E RUBINO M SNC IN LIQUIDAZIONE" e
# "HELIOS GROUP S.R.L. IN LIQUIDAZIONE" sono ragioni sociali vere, e una
# regola che cerca la parola ovunque le butterebbe via.
NOME_DOCUMENTO = {
    "ORIGINALE", "ORIGINALI", "COPIA", "COPIE", "ALLEGATO", "ALLEGATI",
    "DOCUMENTO", "DOCUMENTI", "SCANSIONE", "SCANSIONI", "IMMAGINE",
    "CITAZIONE", "TESTIMONIALE", "MEMORIA", "VERBALE", "RELATA", "PROCURA",
    "NOTIFICA", "NOTIFICAZIONE", "SENTENZA", "DECRETO", "ORDINANZA",
    "RICORSO", "ATTO", "ATTI", "FASCICOLO", "UDIENZA", "COMPARSA", "ISTANZA",
    "CONCLUSIONALE", "DEPOSITO", "TRASMISSIONE", "INVIO", "RICHIESTA",
    "SOLLECITO", "PAGAMENTO", "LIQUIDAZIONE", "COMUNICAZIONE", "AVVISO",
    "AVVERSO", "OGGETTO", "TRASMETTO", "TRASMETTESI", "SEGUE", "VEDI",
    "CARTELLA", "CARTELLE", "PRECETTO", "INTIMAZIONE", "DIFFIDA",
}

def perche_non_e_un_nome(v):
    """Se restituisce un motivo, quello non e' un nominativo ma una frase.
    Meglio nessun nome che uno sbagliato: un nome mancante si vede, uno
    sbagliato entra in archivio e nessuno se ne accorge piu'."""
    t = [p for p in re.split(r"[^A-Z\']+",
         unicodedata.normalize("NFKD", testo(v)).encode("ascii", "ignore").decode().upper())
         if p]
    if not t:
        return None
    if t[0] in NOME_APERTURA:
        return "comincia con \u00ab%s\u00bb" % t[0]
    # "ORIGINALE", "citazione testimoniale": e' il nome del documento finito
    # nel campo del ricorrente. Solo in testa: "SNC IN LIQUIDAZIONE" resta.
    if t[0] in NOME_DOCUMENTO:
        return "comincia con \u00ab%s\u00bb: \u00e8 il nome di un atto, non di una persona" % t[0]
    for p in t:
        if p in NOME_RUOLI:
            return "contiene \u00ab%s\u00bb, che in un nome non c\u2019\u00e8" % p
    if len(t) >= 4:
        for p in t[1:]:
            if p in NOME_INTERNE:
                return "\u00e8 una frase di %d parole, non un nome" % len(t)
    return None

# "QUARTARONE ROSARIA / INPS", "Piazza Claudio contro Inps": il nome della
# causa. E' l'unica forma libera dell'oggetto che si possa leggere senza
# sbagliare - provata su 1.029 oggetti di cui il nominativo era gia' noto:
# otto estratti, otto giusti, nessun falso.
#
# Le altre forme che sembravano ovvie sono state buttate dopo la prova:
# "nome in testa seguito da un numero" leggeva COMUNICAZIONE da
# "COMUNICAZIONE 3267/2024/LAV" - 834 errori su 893 estrazioni.
RE_CAUSA_CONTRO = re.compile(
    r"(?<![A-Za-z])"
    r"([A-Z][A-Za-z']{2,}(?:\s+[A-Z][A-Za-z']+){0,3})"
    r"\s*(?:/|\bc(?:ontro|\.|/)\s*)\s*"
    r"(?:INPS|I\.?\s?N\.?\s?P\.?\s?S)", re.I)

NOME_CAUSA_STOP = {
    "SENTENZA", "COMUNICAZIONE", "RICORSO", "NOTIFICA", "NOTIFICAZIONE",
    "CAUSA", "PRATICA", "OGGETTO", "INVIO", "RISPOSTA", "TRASMISSIONE",
    "PROCEDIMENTO", "GIUDIZIO", "APPELLO", "DECRETO", "ORDINANZA", "ATTO",
    "MEMORIA", "FASCICOLO", "UDIENZA", "SPESE", "PAGAMENTO", "PER", "DALL",
    "DELL", "NELL", "CONTESTUALMENTE", "OPPOSIZIONE", "IMPUGNAZIONE",
}

def _chiave_nome(n):
    t = unicodedata.normalize("NFKD", testo(n).upper())
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = re.sub(r"\+\s*\d+.*$", "", t)
    return frozenset(x for x in re.split(r"[^A-Z]+", t) if len(x) >= 2)

def nome_a_tre_giri(letture):
    """Il nome se almeno due giri (fonti diverse) lo leggono uguale, parola per
    parola; altrimenti None. letture = [(giro, nome), ...]."""
    per = {}
    for giro, nome in letture:
        k = _chiave_nome(nome)
        if len(k) < 2:
            continue
        per.setdefault(k, []).append((giro, nome))
    for k, l in per.items():
        if len({g for g, _ in l}) >= 2:
            return l[0][1]
    return None

def nome_dalla_causa(t):
    """Il nominativo dal nome della causa: "TIZIO / INPS", "Tizio contro Inps"."""
    m = RE_CAUSA_CONTRO.search(testo(t))
    if not m:
        return None
    v = re.sub(r"\s+", " ", m.group(1)).strip(" .-'").upper()
    par = [p for p in re.split(r"[^A-Z']+", v) if len(p) > 1]
    # La ricerca ignora maiuscole e minuscole, quindi puo' agganciare la
    # parola che precede il nome: in "sentenza GALLO TERESA/INPS" prendeva
    # "sentenza GALLO TERESA". Le parole di rito si tolgono davanti invece di
    # far fallire tutta l'estrazione.
    while par and par[0] in NOME_CAUSA_STOP:
        par.pop(0)
    par = [p for p in par if p not in NOME_CAUSA_STOP]
    v = " ".join(par)
    if len(v) < 5 or e_nome_interno(v) or perche_non_e_un_nome(v):
        return None
    return _ripulisci_nome_persona(v)

# "Notificazione ai sensi della Legge N. 53 del 1994 - [ALE] 233 - Bellomo
# Francesco - Decreto di OMOLOGA -": l'avvocato scrive in oggetto il nome
# della parte e l'atto notificato, a pezzi separati da trattini.
RE_NOTIFICA_53 = re.compile(r"\b53\s*(?:/|del)\s*(?:19)?94\b\s*[-\u2013]?(.*)$", re.I)
PAROLE_ATTO = re.compile(r"\b(decreto|ricorso|sentenza|ordinanza|atto|omologa|precetto|citazione"
                         r"|appello|verbale|memoria|istanza|notific\w*|avviso|relata|ingiunzion\w*)\b", re.I)

def dati_da_notifica_53(t):
    """(nome, atto) dall'oggetto di una notifica dell'avvocato, o (None, None)."""
    m = RE_NOTIFICA_53.search(testo(t))
    if not m:
        return None, None
    pezzi = [x.strip(" .,;:") for x in re.split(r"\s+[-\u2013]\s+|\s+[-\u2013]$|^[-\u2013]\s+", m.group(1))]
    pezzi = [x for x in pezzi if x]
    nome = atto = None
    for x in pezzi:
        if re.fullmatch(r"(?:\[[^\]]*\]\s*)?[\d/ ]*", x):
            continue                              # "[ALE] 233": sigla e numero dello studio
        x = re.sub(r"^\[[^\]]*\]\s*\d*\s*", "", x).strip()
        if not x:
            continue
        if PAROLE_ATTO.search(x):
            atto = atto or x
        elif not nome and re.fullmatch(r"[A-Za-z\u00c0-\u00ff'\u2019. ]{4,60}", x) and len(x.split()) >= 2:
            nome = _ripulisci_nome_persona(x)
    return nome, atto

def nome_da_notifica_53(t):
    return dati_da_notifica_53(t)[0]

def nome_da_pratica(t):
    """'(Sede: CATANIA - Pratica: 3689/25/CO/2 - YE SHUFENG)': il nominativo
    e' li' in chiaro, anche quando il PDF allegato non lo contiene."""
    m = RE_PRATICA_NOME.search(testo(t))
    return _ripulisci_nome_persona(m.group(1)) if m else None

# fra l'ufficio e la citta' possono esserci parole di contorno:
# "TRIBUNALE ORDINARIO DI SIRACUSA", "Corte di Appello di Catania"
RE_UFFICIO = re.compile(
    r"\b(tribunale|corte\s*d?\s*[i'\u2019]?\s*appello|corteappello|corte\s+suprema|cassazione)\b"
    r"((?:\s+(?:ordinari[oa]|civile|del\s+lavoro|sezione\s+lavoro|di|d[i'\u2019]?))*)"
    r"\s+([A-Z][A-Za-z\u00c0-\u00ff']{3,20})", re.I)
RE_UFFICIO_PEC = re.compile(
    r"\b(tribunale|corteappello|corte\.appello|cortedappello)[._-]?([a-z]{3,20})@", re.I)

def sede_e_grado(testi, mittente=""):
    """Lo stesso numero di ruolo esiste a Siracusa e a Catania: senza la sede
    l'aggancio puo' portare a un fascicolo che non c'entra. La provenienza si
    legge dall'indirizzo del mittente, dal modulo, dall'XML o dalla sigla (CT)."""
    unito = " ".join(testo(x) for x in (testi if isinstance(testi, (list, tuple)) else [testi]))
    mitt = testo(mittente)

    # (CT) vale SEMPRE Catania (confermato dall'utente il 29/09/2026). Prima
    # stava in fondo: bastava che il testo citasse il "Tribunale di Siracusa"
    # della sentenza di primo grado e l'appello passava per Siracusa.
    if re.search(r"\(\s*CT\s*\)", unito):
        return {"sede": "CATANIA", "grado": "appello"}

    m = RE_UFFICIO_PEC.search(mitt) or RE_UFFICIO_PEC.search(unito)
    if m:
        uff, citta = m.group(1).lower(), m.group(2).upper()
        return {"sede": citta,
                "grado": "appello" if "appello" in uff or "corte" in uff else "primo"}

    m = RE_UFFICIO.search(unito)
    if m:
        uff, citta = m.group(1).lower(), m.group(3).upper()
        if citta not in ("DI", "DEL", "DELLA", "ORDINARIO", "CIVILE", "LAVORO", "SEZIONE"):
            return {"sede": citta,
                    "grado": "cassazione" if "cassaz" in uff or "suprema" in uff
                             else ("appello" if "appello" in uff else "primo")}

    return {"sede": None, "grado": None}

# --------------------------------------------------------------------------
# NOME DEFINITIVO DELLA CARTELLA: RICORRENTE UD ggmmaa
# --------------------------------------------------------------------------
MESI = {"gennaio":1,"febbraio":2,"marzo":3,"aprile":4,"maggio":5,"giugno":6,
        "luglio":7,"agosto":8,"settembre":9,"ottobre":10,"novembre":11,"dicembre":12}

RE_UD = [
    re.compile(r"udienza[^\n]{0,60}?(\d{1,2})\s*[/.\-]\s*(\d{1,2})\s*[/.\-]\s*(\d{2,4})", re.I),
    re.compile(r"udienza[^\n]{0,60}?(\d{1,2})\s+(" + "|".join(MESI) + r")\s+(\d{4})", re.I),
    re.compile(r"comparire[^\n]{0,120}?(\d{1,2})\s*[/.\-]\s*(\d{1,2})\s*[/.\-]\s*(\d{2,4})", re.I),
    re.compile(r"comparire[^\n]{0,120}?(\d{1,2})\s+(" + "|".join(MESI) + r")\s+(\d{4})", re.I),
    re.compile(r"fissa[^\n]{0,80}?(\d{1,2})\s*[/.\-]\s*(\d{1,2})\s*[/.\-]\s*(\d{2,4})", re.I),
]

# NB: niente \s nelle classi, altrimenti il nome scavalca l'a capo e si porta
# dietro la riga successiva ("CAMPISI CORRADINA UDIENZA DEL")
NOME = r"[A-Z\u00c0-\u00dd][^\n\r,;]{2,58}"
RE_RIC = [
    re.compile(r"nell['\u2019]interesse\s+d[ie]\s+(" + NOME + ")", re.I),
    re.compile(r"promosso\s+da\s+(" + NOME + r")\s+(?:contro|nei\s+confronti|c/|c\.)", re.I),
    re.compile(r"ricorso\s+d[ie]\s+(" + NOME + r")\s+(?:contro|nei\s+confronti|c/)", re.I),
    # atti di cancelleria: "TRA <nome> E INPS", "tra <nome> (ricorrente)"
    re.compile(r"^\s*\bTRA\b\s*[:\-]?\s*\n?\s*(" + NOME + ")", re.I | re.M),
    re.compile(r"\btra\s+(" + NOME + r")\s*\(?\s*ricorrente", re.I),
    re.compile(r"\bcausa\s*[:\-]\s*(" + NOME + ")", re.I),
    re.compile(r"\boggetto\s*[:\-][^\n]{0,30}?\bd[ie]\s+(" + NOME + ")", re.I),
    re.compile(r"\bsig(?:\.ra|\.r[ae]|nor[ae])?\.?\s+(" + NOME + ")", re.I),
    re.compile(r"ricorrente\s*[:\-]?\s*(" + NOME + ")", re.I),
    re.compile(r"^\s*([A-Z\u00c0-\u00dd][^\n\r/\\]{3,49})\s*(?:[/\\]|\bc(?:ontro)?\.?\s*)\s*I\.?N\.?P\.?S", re.I | re.M),
]

# NB: "lo", "la", "di", "de" NON vanno qui: sono particelle di cognomi veri
# (LO PRESTI, LA ROSA, DI MAURO, DE LUCA). Le formule di contorno le taglia RE_DOPO_NOME.
FERMA = {"elettivamente","domiciliato","domiciliata","rappresentato","rappresentata",
         "difeso","difesa","contro","confronti","residente","via","viale","piazza",
         "nato","nata","presso","giusta","studio","pec","avv","avvocato","inps",
         "istituto","tribunale","sezione","corte","giudice","codice","fiscale"}

# formule che nei ricorsi seguono il nome: il nome finisce dove comincia una di queste
RE_DOPO_NOME = re.compile(
    r"\b(?:di\s+anni|nat[oa]\s+(?:a|il)|nat[oa]\b|c\.?\s?f\.?|cod(?:ice)?\.?\s*fisc"
    r"|p\.?\s?iva|residente|domicili|elettivamente|rappresentat|difes|in\s+persona"
    r"|contro|nei\s+confronti|con\s+sede|corrente\s+in"
    # formule dei verbali di udienza: il nome finisce dove comincia il verbale
    r"|nessuno|non\s+\u00e8|comparso|compars[ae]|present[ei]|assent[ei]|udienza|verbale"
    r"|il\s+giudice|cancellier|procurator|si\s+d\u00e0\s+atto|a\s+scioglimento)\b", re.I)

NOMI_INTERNI = set()

def e_nome_interno(v):
    """Il ricorrente non puo' essere un avvocato dell'Istituto, un referente
    di una UO o un collega dell'ufficio: quelli sono chi SCRIVE la PEC, non
    chi ricorre. Confondere le due cose riempie l'archivio di fascicoli
    intestati ai colleghi."""
    if not NOMI_INTERNI:
        return False
    n = _compatta(v)
    if not n:
        return False
    for x in NOMI_INTERNI:
        c = _compatta(x)
        if not c:
            continue
        if n == c:
            return True
        # il solo cognome basta, in un verso o nell'altro: "SPADONI" per
        # "SPADONI ERCOLE", e "SPADONI ERCOLE" per "SPADONI"
        parti = [_compatta(p) for p in x.upper().split() if len(p) > 3]
        if not parti:
            continue
        cogn = parti[0]
        if len(cogn) >= 6 and (n == cogn or c.startswith(n) or n.startswith(cogn)):
            return True
    return False

def _ripulisci_nome_persona(v):
    v = re.sub(r"\s+", " ", testo(v)).strip(" .,;:-\u2013")
    v = re.sub(r"^(?:causa|oggetto|tra|ricorso|proc(?:edimento)?|sig(?:\.ra)?|nel)\b[\s:\-]*", "",
               v, flags=re.I).strip()
    m = RE_DOPO_NOME.search(v)
    if m:
        v = v[:m.start()].strip(" .,;:-")
    v = re.split(r"\d", v)[0].strip(" .,;:-")   # "Alfredo 67" -> "Alfredo"
    v = re.split(r"[/\\(\[]", v)[0].strip(" .,;:-")   # "X / INPS", "X (ricorrente)"
    v = re.sub(r"\s+c(?:ontro)?\.?\s*$", "", v, flags=re.I).strip(" .,;:-")
    fuori = []
    for i, parola in enumerate(v.split()):
        pulita = re.sub(r"[^a-z\u00e0-\u00ff]", "", parola.lower())
        if i > 0 and pulita in FERMA:      # la prima parola non si taglia mai
            break
        fuori.append(parola)
    v = " ".join(fuori).strip(" .,;:-")
    # congiunzioni e iniziali orfane in coda: "ALECCI ROSARIA E" -> "ALECCI ROSARIA"
    v = re.sub(r"\s+(?:e|ed|et|&|con|nonch[e\u00e9])\s*$", "", v, flags=re.I).strip(" .,;:-")
    if len(v) < 3 or len(v) > 55:
        return None
    secco = re.sub(r"[^A-Z]", "", v.upper())      # "I.N.P.S." -> "INPS"
    if any(x in secco for x in ("TRIBUNALE", "INPS", "ISTITUTONAZIONALE", "STUDIO",
                                "AVVOCAT", "CORTEDAPPELLO", "CANCELLERIA")):
        return None
    if len(v.split()) > 5:        # un nominativo non ha sei parole: e' testo inglobato
        return None
    # qualifiche e oggetti del giudizio catturati al posto del nome
    NON_NOMI = {"APPELLANTE","APPELLATO","RICORRENTE","RESISTENTE","CONVENUTO","ATTORE",
                "OPPONENTE","OPPOSTO","DEBITORE","CREDITORE","CONTROPARTE","PARTE",
                "LICENZIAMENTO","RISARCIMENTO","PENSIONE","INVALIDITA","ACCOMPAGNAMENTO",
                "CONTRIBUTI","INDEBITO","SENTENZA","RICORSO","DECRETO","UDIENZA",
                "COMUNICAZIONE","NOTIFICA","OGGETTO","CAUSA","PROCEDIMENTO","FASCICOLO",
                "ADDEBITO","AVVISODIADDEBITO","CARTELLAESATTORIALE","INTIMAZIONE",
                "DINIEGO","ACCERTAMENTO","PRESTAZIONE","OPPOSIZIONE","IMPUGNAZIONE",
                "COSTITUZIONE","MEMORIA","ATTO","VERBALE","RELATA","PROCURA",
                "ALLEGATI","ALLEGATO","TRASMISSIONE","RICHIESTA","SOLLECITO",
                "PAGAMENTO","LIQUIDAZIONE","RATEAZIONE","SGRAVIO","DURC","ISTRUTTORIA",
                "SISCO","RIFSISCO","RIF","RIFERIMENTO","PRATICA","PRATICHE","NPRAT",
                "PROT","PROTOCOLLO","FASCICOLO","FASCICOLI","RICORSO","RICORSI",
                "ATP","ATPEXART","EXART","CTU","NASPI","RDC","ADI","TFR","DURC",
                "INVCIV","SERVIZIO","SEGRETERIA","PROTOCOLLO","DIREZIONE"}
    if secco in NON_NOMI:
        return None
    if e_nome_interno(v):
        return None
    # un nominativo non comincia con la parola che descrive l'atto
    if re.match(r"^(?:AVVISO|CARTELLA|ATTO|RICORSO|DECRETO|SENTENZA|VERBALE|MEMORIA|"
                r"COMUNICAZIONE|RICHIESTA|TRASMISSIONE|ALLEGAT|OPPOSIZIONE|IMPUGNAZIONE|"
                r"ADDEBITO|ISTANZA|DOMANDA|PRATICA|FASCICOLO)\b", v.upper()):
        return None
    # una parola sola con desinenza da sostantivo astratto non e' un cognome.
    # -SIONE/-ZIONE e non il generico -IONE: MARCHIONE e CAMPIONE sono cognomi veri.
    if len(v.split()) == 1 and re.search(r"(ANTE|MENTO|[SZ]IONE)$", secco):
        return None
    if re.match(r"^(avv|dott|dr|prof|sig|geom|ing|rag)\.?\b", v, re.I):
        return None          # e' una qualifica, non il ricorrente
    return v.upper()

NOMEFILE_STOP = set("""FATTURA FATTURE RICORSO RICORSI COPIA TELEMATICA DECRETO DECRETI
SENTENZA SENTENZE ATTO ATTI NOTIFICA NOTIFICAZIONE RELATA PROCURA LITI VERBALE VERBALI
ORDINANZA COMUNICAZIONE LETTERA ACQUIESCENZA ALLEGATO ALLEGATI POSTACERT DATICERT IMAGE
INDICEBUSTA DOC PDF EML XML P7M ZIP GRADO SEZ SEZIONE TRIB TRIBUNALE CORTE APPELLO INPS
FAVOREVOLE SFAVOREVOLE PER CON RGE UDIENZA SPESE LEGALI LEGALE PAGAMENTO BONIFICO
ISTRUTTORIA RICHIESTA TRASM INVIO POSTA ELETTRONICA SIRACUSA CATANIA RAGUSA NUOVO PRATICA
SEDE ESITO CITAZIONE MEMORIA DIFENSIVA CONCLUSIONALE NOTA COMPARSA ISTANZA OPPOSIZIONE
MARZO GENNAIO FEBBRAIO APRILE MAGGIO GIUGNO LUGLIO AGOSTO SETTEMBRE OTTOBRE NOVEMBRE
DICEMBRE ALLE ALLA AGLI DEGLI DELLE DELLA DALLA DALLE FAV SIG SIGRA SIGNOR SIGNORA
ESECUZIONE PRECETTO PIGNORAMENTO SURROGA CTU PERIZIA CONSULENZA
ALL ALLEG UL INGIUNTIVO ANAGRAFICA AZIENDA VISURA SITUAZIONE DEBITORIA DIFFIDE DIFFIDA
EREDI SCARICATO SCARICATA DOCUMENTO DOCUMENTI ESTRATTO CONTO VERIFICA VERIFICHE ELENCO
PROSPETTO RIEPILOGO CERTIFICATO CERTIFICAZIONE MODULO ISTANZE DEPOSITO RICEVUTA""".split())
NOMEFILE_PART = {"LO", "LA", "LI", "LE", "DE", "DI", "DA", "DEL", "DELLO", "SAN", "SANTA"}

def _pezzi_nomefile(n):
    """I frammenti di un nome di file che potrebbero essere un nominativo."""
    b = re.sub(r"\.(pdf|p7m|eml|xml|zip|docx?|png|jpe?g|tiff?)$", "", testo(n), flags=re.I)
    b = re.sub(r"\.(pdf|p7m)$", "", b, flags=re.I)
    b = unicodedata.normalize("NFKD", b).encode("ascii", "ignore").decode().upper()
    b = re.sub(r"\b[A-Z]\.", "", b)
    b = re.sub(r"\d+[\s/._\-]*\d*[\s/._\-]*\d*", " ", b)
    b = re.sub(r"[^A-Z\' ]+", " ", b)
    par = [x for x in b.split()
           if x not in NOMEFILE_STOP and re.search(r"[AEIOU]", x)]
    tenuti = []
    for i, x in enumerate(par):
        if x in NOMEFILE_PART:
            if i + 1 < len(par) and par[i + 1] not in NOMEFILE_PART and len(par[i + 1]) > 2:
                tenuti.append(x)
        elif len(x) > 2:
            tenuti.append(x)
    fuori, salta = [], False
    for i, x in enumerate(tenuti):
        if salta:
            salta = False
            continue
        if x in NOMEFILE_PART and i + 1 < len(tenuti):
            fuori.append(x + " " + tenuti[i + 1])
            salta = True
        else:
            fuori.append(x)
    return fuori

def nome_da_nomefile(nomi):
    """Il nominativo sta spesso nel nome del file. E' l'ultima fonte, dopo il
    modulo, l'XML, l'oggetto e il testo dei documenti.

    Il criterio non e' 'il primo file che dice qualcosa': su
        ALL 2 Anagrafica Azienda.pdf, ALL 3 visura bellofiore giovanni.pdf,
        ALL 5 diffide eredi bellofiore.pdf, decreto ingiuntivo Bellofiore.pdf
    il primo dava 'ALL ANAGRAFICA AZIENDA'. Il nome giusto e' quello che
    RITORNA: gli allegati di un ricorso portano tutti lo stesso cognome, le
    parole di contorno no."""
    nomi = nomi if isinstance(nomi, (list, tuple)) else [nomi]
    per_file = [_pezzi_nomefile(n) for n in nomi]
    freq = {}
    for l in per_file:
        for x in set(l):
            freq[x] = freq.get(x, 0) + 1
    ripetuti = [k for k, v in freq.items() if v > 1]
    if ripetuti:
        best = max(ripetuti, key=lambda k: (freq[k], len(k)))
        for l in per_file:
            if best in l:
                i = l.index(best)
                v = (l[i] + " " + l[i + 1]) if i + 1 < len(l) and l[i + 1] not in ripetuti else best
                if len(v) >= 4 and not e_nome_interno(v) and not perche_non_e_un_nome(v):
                    return _ripulisci_nome_persona(v)
        return None
    for l in per_file:
        if not l:
            continue
        v = " ".join(l[:3])
        if len(v) >= 4 and not e_nome_interno(v) and not perche_non_e_un_nome(v):
            return _ripulisci_nome_persona(v)
    return None

def testo_dei_documenti(cartella, max_pagine=4, max_file=6, ricorsivo=False):
    """Legge il testo dei PDF. Attenzione all'ordine e alle pagine: il giudice
    firma in fondo ai provvedimenti che emette, non nel ricorso — che gli e'
    rivolto quando ancora non si sa chi sara'. Quindi prima i provvedimenti,
    e di ognuno anche le ULTIME pagine, dove sta la sottoscrizione."""
    try:
        from pypdf import PdfReader
    except ImportError:
        try:
            from PyPDF2 import PdfReader
        except ImportError:
            return "", "manca pypdf: nome cartella non ricavabile dai documenti"

    def priorita(p):
        n = p.name.lower()
        if re.search(r"decret|fissaz|dfu|ordinanz|sentenz|verbal|provvedim", n):
            return 0          # li' c'e' la firma del giudice
        if re.search(r"ricors|atto|citazion", n):
            return 1          # li' c'e' il ricorrente e l'oggetto
        return 2

    # anche i .p7m: gli atti del processo telematico sono firmati, e dentro
    # la busta c'e' il PDF. Cercare solo *.pdf significa non vederli.
    tutti = cartella.rglob("*") if ricorsivo else cartella.glob("*")
    trova = [p for p in tutti if p.is_file() and
             re.search(r"\.pdf(\.p7m)?$", p.name, re.I)]
    ordine = sorted(trova, key=lambda p: (priorita(p), p.name))
    pezzi, letti = [], 0
    for p in ordine:
        if letti >= max_file:
            break
        try:
            sorgente = str(p)
            if p.suffix.lower() == ".p7m":
                dentro = apri_p7m(p.read_bytes())
                if not dentro:
                    continue
                import tempfile as _tf
                tmp = Path(_tf.gettempdir()) / ("_p7m_%d.pdf" % abs(hash(str(p))))
                tmp.write_bytes(dentro)
                sorgente = str(tmp)
            r = PdfReader(sorgente)
            n = len(r.pages)
            # le prime pagine per l'oggetto e le parti, le ultime per la firma
            indici = list(range(min(max_pagine, n)))
            for i in range(max(0, n - 2), n):
                if i not in indici:
                    indici.append(i)
            for i in indici:
                pezzi.append(r.pages[i].extract_text() or "")
            letti += 1
        except Exception:
            continue
    return "\n".join(pezzi), None

# La data del RINVIO, dal modulo di cancelleria, dov'e' scritta in forma
# fissa: "RINVIATO AD ALTRA UDIENZA DI DISCUSSIONE IL 01/12/2026 09:00".
# trova_udienza cerca genericamente la prima data dopo la parola "udienza" in
# TUTTO il testo - modulo, corpo, PDF - e puo' vincere la riga "Notificato
# alla PEC / in cancelleria il 26/05/2026", che e' la data della notifica, non
# quella dell'udienza. Su un rinvio significa scrivere nel fascicolo il giorno
# sbagliato, ed e' peggio di non scrivere niente.
RE_RINVIO_MODULO = re.compile(
    r"rinviat\w*\s+ad?\s+altra\s+udienza[^0-9]{0,60}?"
    r"(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})", re.I)

RE_FISSATA_MODULO = re.compile(
    r"(?:fissat\w*|udienza\s+di\s+comparizione)[^0-9]{0,60}?"
    r"(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})", re.I)

# Le formule del VERBALE, che il giudice detta a udienza e che arrivano
# allegate alla comunicazione di cancelleria. Sono la fonte vera del rinvio:
# il modulo a volte non lo scrive, il verbale sempre.
#   "rinvia all'udienza del 14/11/2026"
#   "si rinvia al 14.11.2026 per la discussione"
#   "nuova udienza 14/11/2026"
#   "rinvio la causa all'udienza del 14 novembre 2026"
RE_RINVIO_VERBALE = [
    re.compile(r"rinvi\w*[^\n]{0,80}?(\d{1,2})\s*[/.\-]\s*(\d{1,2})\s*[/.\-]\s*(\d{2,4})", re.I),
    re.compile(r"rinvi\w*[^\n]{0,80}?(\d{1,2})\s+(" + "|".join(MESI) + r")\s+(\d{4})", re.I),
    re.compile(r"nuova\s+udienza[^0-9]{0,40}?(\d{1,2})\s*[/.\-]\s*(\d{1,2})\s*[/.\-]\s*(\d{2,4})", re.I),
    re.compile(r"udienza\s+del\s+(\d{1,2})\s*[/.\-]\s*(\d{1,2})\s*[/.\-]\s*(\d{2,4})", re.I),
]

def _data_da(m):
    """(gg, mm|nome_mese, aaaa) -> date, o None se non e' una data sensata."""
    g, mese, a = m.group(1), m.group(2), m.group(3)
    mm = MESI.get(str(mese).lower()) if not str(mese).isdigit() else int(mese)
    try:
        a = int(a)
        if a < 100:
            a += 2000
        d = date(a, int(mm), int(g))
    except Exception:
        return None
    return d if 2015 <= d.year <= 2035 else None

def udienza_annunciata(testi, dopo=None):
    """La data NUOVA che l'atto annuncia: il rinvio.

    Due regole, e sono quelle che evitano di scrivere nel fascicolo il giorno
    sbagliato:
      1) si guardano solo le frasi che parlano di rinvio o di fissazione - non
         una qualunque data vicina alla parola 'udienza', che puo' essere
         l'udienza appena celebrata o la data della notifica;
      2) un rinvio va sempre AVANTI: si tengono solo le date successive a
         'dopo' (la data di ricezione della PEC), e fra queste si prende la
         piu' vicina - se il verbale ne cita due, la prima utile e' quella
         fissata.
    Restituisce (data, frase) oppure (None, None)."""
    trovate = []
    for t in testi:
        tt = testo(t)
        if not tt:
            continue
        for rx in [RE_RINVIO_MODULO, RE_FISSATA_MODULO] + RE_RINVIO_VERBALE:
            for m in rx.finditer(tt):
                d = _data_da(m)
                if not d:
                    continue
                if dopo and d <= dopo:
                    continue
                frase = re.sub(r"\s+", " ", tt[max(0, m.start() - 40): m.end() + 20]).strip()
                trovate.append((d, frase))
    if not trovate:
        return None, None
    trovate.sort(key=lambda x: x[0])
    return trovate[0]

def udienza_dal_modulo(*testi):
    """La data dell'udienza come la scrive la cancelleria. None se il modulo
    non parla di udienze: allora vale la ricerca generica."""
    for t in testi:
        for rx in (RE_RINVIO_MODULO, RE_FISSATA_MODULO):
            m = rx.search(testo(t))
            if not m:
                continue
            try:
                d = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except Exception:
                continue
            if 2015 <= d.year <= 2035:
                return d
    return None

def trova_udienza(t):
    for rx in RE_UD:
        m = rx.search(t or "")
        if not m:
            continue
        g, mese, a = m.group(1), m.group(2), m.group(3)
        mm = MESI.get(str(mese).lower()) if not str(mese).isdigit() else int(mese)
        try:
            if not mm or not (1 <= mm <= 12):
                continue
            a = int(a)
            if a < 100:
                a += 2000
            d = date(a, mm, int(g))
        except Exception:
            continue
        if 2015 <= d.year <= 2035:
            return d
    return None

def trova_ricorrente(t, oggetto=""):
    for rx in RE_RIC:
        m = rx.search(t or "")
        if m:
            v = _ripulisci_nome_persona(m.group(1))
            if v:
                return v
    # ripiego sull'oggetto: "... per Lo Bello / INPS", "... - Campisi Corradina"
    o = testo(oggetto)
    m = re.search(r"\bper\s+([A-Z\u00c0-\u00dd][A-Za-z\u00c0-\u00ff'\u2019\s\.\-]{2,40})\s*[/\\]\s*INPS", o, re.I)
    if m:
        v = _ripulisci_nome_persona(m.group(1))
        if v:
            return v
    m = re.search(r"[\-\u2013]\s*([A-Z\u00c0-\u00dd][A-Z\u00c0-\u00dd'\u2019\s]{4,40})\s*$", o)
    if m:
        return _ripulisci_nome_persona(m.group(1))
    return None

def giudice_canonico(nome):
    n = unicodedata.normalize("NFKD", testo(nome)).encode("ascii", "ignore").decode().upper().strip()
    n = re.sub(r"^(DR\.?SSA|DOTT\.?SSA|DR\.?|DOTT\.?|AVV\.?|SIG\.?RA|SIG\.?)\s*", "", n)
    n = re.sub(r"[^A-Z' ]", "", n)
    n = re.sub(r"\s+", " ", n).strip()
    if len(n) > 3 and n[0] == n[1]:
        n = n[1:]
    return n

def giudice_seguito(nome, cfg, grado=None, sede=None):
    """La ripartizione per giudice vale in primo grado, dove ogni operatore ha
    i suoi. In appello i giudici non sono assegnati a nessuno: li' il fascicolo
    si riconosce dal numero di ruolo, e non si esclude niente per il giudice."""
    if grado == "appello" or (sede and giudice_canonico(sede).startswith("CATANIA")):
        return None
    seguiti = [giudice_canonico(x) for x in (cfg.get("giudici_seguiti") or []) if x]
    if not seguiti:
        return None                    # nessuna ripartizione indicata: non decido
    g = giudice_canonico(nome)
    if not g:
        return None                    # giudice ignoto: meglio lavorarla

    # Il confronto NON deve dipendere dall'ordine. Il nome del giudice arriva
    # da fonti diverse - il modulo di cancelleria lo scrive "PEDALINO
    # GIOVANNA", il testo di una sentenza "Il Giudice dott.ssa Giovanna
    # Pedalino" - e confrontando le prime lettere nell'ordine in cui capitano,
    # "GIOVANNA PEDALINO" non risultava fra i seguiti: la PEC finiva fra i
    # giudici di altri e spariva dalla lavorazione.
    def parole(v):
        return sorted(p for p in re.split(r"[^A-Z]+", v) if len(p) > 2)

    pg = parole(g)
    if not pg:
        return None
    for x in seguiti:
        px = parole(x)
        if not px:
            continue
        # Due nomi completi sono la stessa persona se hanno le STESSE parole,
        # in qualunque ordine. Una parola sola in comune non basta: "FILIPPO
        # RUSSO" e "FAVALE FILIPPO" condividono il nome di battesimo e non
        # sono la stessa persona.
        if len(pg) >= 2 and len(px) >= 2:
            if set(pg) <= set(px) or set(px) <= set(pg):
                return True
            continue
        # Un nome solo - capita quando dal documento esce il solo cognome -
        # vale se e' il COGNOME di un giudice seguito. Nella configurazione i
        # nomi si scrivono cognome prima, quindi e' la prima parola.
        solo = (pg if len(pg) == 1 else px)[0]
        cognome = [p for p in re.split(r"[^A-Z]+", giudice_canonico(x)) if len(p) > 2]
        if cognome and solo == cognome[0] and len(solo) >= 4:
            return True
    return False

def nome_definitivo(ricorrente, udienza, rif=None, quando=None, tipo=None, oggetto=""):
    """Convenzione piena dell'ufficio: RICORRENTE UD ggmmaa.
    Se manca l'udienza si usa comunque il ricorrente, con il riferimento e la
    data: una cartella con un nome parziale si trova, una senza nome no."""
    if ricorrente and udienza:
        return pulisci_nome(f"{ricorrente} UD {udienza:%d%m%y}", 120), True
    if not ricorrente:
        return None, False
    pezzi = [ricorrente]
    r = rif or {}
    for eti, ch in (("SISCO", "sisco"), ("RG", "rg"), ("sent", "sentenza"),
                    ("com", "comunicazione"), ("spese", "spese_legali")):
        if r.get(ch):
            pezzi.append(f"{eti} {str(r[ch]).replace('/', '-')}")
            break
    if quando:
        pezzi.append(str(quando)[:10])
    if tipo and tipo not in ("da_classificare", "da_leggere"):
        pezzi.append(tipo.replace("_", " "))
    return pulisci_nome(" - ".join(pezzi), 120), False

def unisci_cartelle(origine, destinazione, esito):
    """Due PEC diverse sullo stesso fascicolo (ricorso e decreto arrivati
    separatamente) devono finire nella stessa cartella: il nome
    RICORRENTE UD ggmmaa e' l'identita' del fascicolo, non del messaggio."""
    spostati, gia = [], 0
    impronte = {}
    for x in destinazione.iterdir():
        if x.is_file():
            try:
                impronte[hashlib.sha256(x.read_bytes()).hexdigest()] = x.name
            except Exception:
                pass
    for x in sorted(origine.iterdir()):
        if not x.is_file() or x.name in ("scheda_pec.json", "testo_pec.txt"):
            continue
        try:
            h = hashlib.sha256(x.read_bytes()).hexdigest()
        except Exception:
            h = None
        if h and h in impronte:
            gia += 1
            continue
        nuovo = destinazione / x.name
        n = 2
        while nuovo.exists():
            nuovo = destinazione / f"{x.stem} ({n}){x.suffix}"
            n += 1
        shutil.move(str(x), str(nuovo))
        spostati.append(nuovo.name)
        if h:
            impronte[h] = nuovo.name
    if gia:
        esito["note"].append(f"{gia} documenti gia' presenti nella cartella del fascicolo")
    return spostati

# --------------------------------------------------------------------------
# FASCICOLI GIA' APERTI: le comunicazioni vanno in coda, non fanno cartella nuova
# --------------------------------------------------------------------------
FASC = {"per_sisco": {}, "per_rg": {}, "per_chiave": {}, "cartella_rg": {},
        "cartella_sisco": {}, "radice": None, "quante": 0, "formato_mese": None,
        "puo_creare": False, "rg_ambigui": {}}

MESI_SIGLA = ["GEN","FEB","MAR","APR","MAG","GIU","LUG","AGO","SET","OTT","NOV","DIC"]

def _compatta(s):
    return re.sub(r"[^A-Z0-9]", "",
                  unicodedata.normalize("NFKD", testo(s)).encode("ascii", "ignore").decode().upper())

def _indicizza_rg(chiave, voce):
    """Il numero di ruolo NON e' univoco: lo stesso 1438/2026 puo' esistere a
    Siracusa e a Catania, perche' ogni tribunale numera per conto suo. Finche'
    nell'archivio ce n'e' uno solo il ruolo decide bene; quando ne compaiono
    due il numero smette di poter decidere, e resta scritto qui. Da quel
    momento per quei fascicoli aggancia solo il numero di pratica, che il
    suffisso (CO, AO, CV, AI, IC) rende unico."""
    if not chiave:
        return
    gia = FASC["per_rg"].get(chiave)
    if gia is None:
        FASC["per_rg"][chiave] = voce
        return
    if gia.get("id") == voce.get("id"):
        return
    q = FASC["rg_ambigui"].setdefault(
        chiave, [testo(gia.get("ricorrente")) or testo(gia.get("id"))])
    n = testo(voce.get("ricorrente")) or testo(voce.get("id"))
    if n and n not in q:
        q.append(n)

def _chiave_fascicolo(nome, ud):
    """La stessa chiave che usa il registro: NOMECOMPATTO|ggmmaa."""
    n = _compatta(nome)
    return f"{n}|{ud}" if n and ud else None

def _legge_nome_cartella(nome):
    m = re.search(r"(\d{6})\s*$", testo(nome).strip())
    if not m:
        return None
    testa = re.sub(r"[\s_\-.]*UD[\s_\-.]*$", "", nome[:m.start()], flags=re.I)
    parte = _compatta(testa)
    return f"{parte}|{m.group(1)}" if parte else None

def _impara_formato_mese(radice):
    """Le cartelle madri esistenti dicono come si chiamano: '0726 LUG',
    'LUG26', '2026-07'. Meglio riusare la forma gia' in uso che imporne una."""
    campioni = [d.name for d in radice.iterdir() if d.is_dir()][:200]
    prove = [
        ("MMAA MES", re.compile(r"^\d{4}\s+[A-Z]{3}$", re.I)),
        ("MESAA",    re.compile(r"^[A-Z]{3}\d{2}$", re.I)),
        ("MES AA",   re.compile(r"^[A-Z]{3}\s+\d{2}$", re.I)),
        ("AAAA-MM",  re.compile(r"^\d{4}-\d{2}$")),
    ]
    conta = {}
    for n in campioni:
        for eti, rx in prove:
            if rx.match(n.strip()):
                conta[eti] = conta.get(eti, 0) + 1
    if conta:
        return max(conta.items(), key=lambda x: x[1])[0]
    return None

def nome_mese(d, formato):
    """d = data dell'udienza. Restituisce il nome della cartella madre."""
    sig = MESI_SIGLA[d.month - 1]
    f = formato or "MESAA"
    if f == "MMAA MES":  return f"{d.month:02d}{d.year % 100:02d} {sig}"
    if f == "MES AA":    return f"{sig} {d.year % 100:02d}"
    if f == "AAAA-MM":   return f"{d.year:04d}-{d.month:02d}"
    return f"{sig}{d.year % 100:02d}"

def costruisci_indice_fascicoli(cfg):
    """Tre fonti, in ordine di affidabilita':
       1) indice_cartelle.json prodotto dal registro: ruolo -> cartella, un salto solo
       2) i marcatori .fascicolo.json lasciati nelle cartelle dalle corse precedenti
       3) base.json + nome della cartella, che e' la strada fragile"""
    FASC.update({"per_sisco": {}, "per_rg": {}, "per_chiave": {}, "cartella_rg": {},
                 "cartella_sisco": {}, "radice": None, "quante": 0, "rg_ambigui": {},
                 "formato_mese": None, "puo_creare": bool(cfg.get("puo_creare_fascicoli", True))})

    radice = cfg.get("archivio_fascicoli")
    if not radice and not cfg.get("base_json") and not cfg.get("indice_cartelle"):
        if SIMULAZIONE:
            log("Instradamento spento: e' cosi' che deve essere in simulazione.")
            log("Ogni PEC creera' la sua cartella, senza agganci al registro.")
            return
        err("=" * 68)
        err("Instradamento SPENTO: in config_pec.json mancano i percorsi.")
        err("  archivio_fascicoli -> la cartella dei fascicoli fisici (RICORSI ES)")
        err("  indice_cartelle    -> indice_cartelle.json esportato dal registro")
        err("  base_json          -> base.json del registro (ripiego)")
        err("Le comunicazioni creeranno cartelle nuove invece di aggiornare i fascicoli.")
        err("=" * 68)
        return

    def assoluto(p):
        q = Path(p)
        return q if q.is_absolute() else BASE / q

    # --- radice dell'archivio ---
    if radice:
        p = assoluto(radice)
        if p.exists():
            FASC["radice"] = p
            FASC["formato_mese"] = _impara_formato_mese(p)
            marcatori = 0
            for d in p.rglob("*"):
                if not d.is_dir():
                    continue
                k = _legge_nome_cartella(d.name)
                if k:
                    FASC["per_chiave"].setdefault(k, d)
                mk = d / ".fascicolo.json"
                if mk.exists():
                    m = leggi_json(mk, {})
                    if m.get("rg"):
                        FASC["cartella_rg"].setdefault(re.sub(r"[^0-9]", "", m["rg"]), d)
                        marcatori += 1
                    if m.get("sisco"):
                        FASC["cartella_sisco"].setdefault(_compatta(m["sisco"]), d)
            log(f"{len(FASC['per_chiave'])} cartelle-fascicolo in {p.name}"
                + (f", {marcatori} gia' marcate" if marcatori else "")
                + (f", cartelle madri in formato {FASC['formato_mese']}" if FASC["formato_mese"] else ""))
        else:
            err(f"archivio_fascicoli non trovato: {p}")

    # --- indice del registro: la strada diretta ---
    ind = cfg.get("indice_cartelle")
    if ind:
        p = assoluto(ind)
        dati = leggi_json(p, None)
        if dati and dati.get("voci"):
            diretti = 0
            for v in dati["voci"]:
                voce = {"id": v.get("id"), "ricorrente": v.get("ricorrente"),
                        "sisco": v.get("sisco"), "rg": v.get("rg"),
                        "sede": v.get("sede"), "grado": v.get("grado"),
                        "udienze": sorted(v.get("udienze") or [], reverse=True)}
                if v.get("rg"):
                    _indicizza_rg(re.sub(r"[^0-9]", "", v["rg"]), voce)
                if v.get("sisco"):
                    FASC["per_sisco"].setdefault(_compatta(v["sisco"]), voce)
                if v.get("cartella") and FASC["radice"]:
                    d = FASC["radice"] / v["cartella"]
                    if not d.exists():
                        d = next((x for x in FASC["radice"].rglob("*")
                                  if x.is_dir() and x.name == Path(v["cartella"]).name), None)
                    if d:
                        diretti += 1
                        if v.get("rg"):
                            FASC["cartella_rg"][re.sub(r"[^0-9]", "", v["rg"])] = d
                        if v.get("sisco"):
                            FASC["cartella_sisco"][_compatta(v["sisco"])] = d
            FASC["quante"] = len(dati["voci"])
            log(f"indice del registro: {FASC['quante']} fascicoli, {diretti} con cartella diretta")
            if FASC["rg_ambigui"]:
                log(f"{len(FASC['rg_ambigui'])} numeri di ruolo su piu' fascicoli: "
                    f"per quelli aggancia solo il numero di pratica")
        else:
            err(f"indice_cartelle non leggibile: {p}. Esportalo dal registro con "
                f"«Esporta indice cartelle».")

    # --- base.json: ripiego, ricostruisce il nome ---
    base = cfg.get("base_json")
    if base and not FASC["per_rg"]:
        p = assoluto(base)
        dati = leggi_json(p, None)
        if not dati:
            err(f"base.json non leggibile: {p}")
            return
        ud_per_fasc = {}
        for u in dati.get("udienze", []):
            if u.get("data") and u.get("fascicolo_id"):
                ud_per_fasc.setdefault(u["fascicolo_id"], []).append(u["data"])
        for f in dati.get("fascicoli", []):
            voce = {"id": f["id"], "ricorrente": f.get("ricorrente"),
                    "sisco": f.get("sisco"), "rg": f.get("rg"),
                    "sede": f.get("sede"), "grado": f.get("grado"),
                    "udienze": sorted(ud_per_fasc.get(f["id"], []), reverse=True)}
            if f.get("sisco"):
                FASC["per_sisco"].setdefault(_compatta(f["sisco"]), voce)
            if f.get("rg"):
                _indicizza_rg(re.sub(r"[^0-9]", "", f["rg"]), voce)
            # nome + udienza: e' cosi' che si ritrova la cartella sul disco.
            # Senza questo l'instradamento non funziona, e le comunicazioni
            # creano cartelle nuove invece di entrare nei fascicoli.
            n = _compatta(f.get("ricorrente"))
            if n:
                for d in ud_per_fasc.get(f["id"], []):
                    try:
                        a, m, g = str(d)[:10].split("-")
                    except Exception:
                        continue
                    FASC["per_chiave"].setdefault(f"{n}|{g}{m}{a[2:]}", voce)
                if not ud_per_fasc.get(f["id"]):
                    FASC["per_chiave"].setdefault(f"{n}|", voce)
        FASC["quante"] = len(dati.get("fascicoli", []))
        log(f"{FASC['quante']} fascicoli e {len(FASC['per_chiave'])} nomi letti da {p.name}")
        if FASC["rg_ambigui"]:
            log(f"{len(FASC['rg_ambigui'])} numeri di ruolo su piu' fascicoli: "
                f"per quelli aggancia solo il numero di pratica")

def cartella_del_fascicolo(voce):
    """Dal fascicolo alla cartella. Il confronto sul nome tollera le
    abbreviazioni: 'BUFALINO C' e 'BUFALINO CAROLA' sono lo stesso."""
    if not voce or not voce.get("ricorrente"):
        return None
    atteso = _compatta(voce["ricorrente"])
    for d in voce.get("udienze", []):
        try:
            a, m, g = d.split("-")
            ud = f"{g}{m}{a[2:]}"
        except Exception:
            continue
        k = f"{atteso}|{ud}"
        if k in FASC["per_chiave"]:
            return FASC["per_chiave"][k]
        for k2, cart in FASC["per_chiave"].items():
            nome, ud2 = k2.rsplit("|", 1)
            if ud2 == ud and nome and (nome.startswith(atteso[:max(5, len(nome))])
                                       or atteso.startswith(nome[:max(5, len(atteso))])):
                return cart
    return None

def crea_cartella_fascicolo(voce):
    """Il fascicolo esiste in registro ma non ha ancora una cartella:
    la creo secondo la convenzione, dentro la cartella madre del mese."""
    if not FASC["puo_creare"] or not FASC["radice"] or not voce.get("ricorrente"):
        return None, "creazione disattivata o dati insufficienti"
    date_ud = voce.get("udienze") or []
    if not date_ud:
        return None, "il fascicolo non ha udienze: nome non componibile"
    try:
        a, m, g = date_ud[0].split("-")
        d = date(int(a), int(m), int(g))
    except Exception:
        return None, "data di udienza illeggibile"
    madre = FASC["radice"] / nome_mese(d, FASC["formato_mese"])
    madre.mkdir(parents=True, exist_ok=True)
    cart = madre / pulisci_nome(f"{voce['ricorrente']} UD {d:%d%m%y}", 120)
    era = cart.exists()
    cart.mkdir(exist_ok=True)
    if not era:
        LANCIO.setdefault("fascicoli_creati", []).append(str(cart))
    k = _chiave_fascicolo(voce["ricorrente"], f"{d:%d%m%y}")
    if k:
        FASC["per_chiave"][k] = cart
    return cart, f"creata in {madre.name}"

def marca_cartella(cart, voce, rif):
    """Lascio dentro la cartella il suo numero di ruolo: dalla prossima corsa
    si riconosce da sola, senza passare da nomi e ricostruzioni."""
    try:
        p = cart / ".fascicolo.json"
        m = leggi_json(p, {})
        m.setdefault("rg", (voce or {}).get("rg") or (rif or {}).get("rg"))
        m.setdefault("sisco", (voce or {}).get("sisco") or (rif or {}).get("sisco"))
        m.setdefault("ricorrente", (voce or {}).get("ricorrente"))
        m["aggiornato"] = datetime.now().isoformat(timespec="seconds")
        scrivi_json(p, m)
        if m.get("rg"):
            FASC["cartella_rg"][re.sub(r"[^0-9]", "", str(m["rg"]))] = cart
        if m.get("sisco"):
            FASC["cartella_sisco"][_compatta(m["sisco"])] = cart
    except Exception:
        pass

def _sede_compatibile(voce, r):
    """Lo stesso ruolo esiste a Siracusa e a Catania: se le sedi sono note e
    diverse, NON e' lo stesso fascicolo. Meglio non agganciare che agganciare
    male, perche' un aggancio sbagliato si presenta come corretto."""
    sede_pec = (r or {}).get("sede")
    sede_f = (voce or {}).get("sede")
    if not sede_pec or not sede_f:
        return True                      # una delle due non si sa: non posso escludere
    return _compatta(sede_pec) == _compatta(sede_f)

def trova_fascicolo(rif, ricorrente=None, udienza=None):
    """(cartella, come_trovata, creata). Il NUMERO DI PRATICA per primo: il
    suffisso (CO, AO, CV, AI, IC) lo rende unico per fascicolo. Il numero di
    ruolo viene dopo, perche' ogni tribunale numera per conto suo e lo stesso
    943/2026 puo' stare su due fascicoli in due sedi: quando nell'archivio ce
    ne sono davvero due, il ruolo non aggancia piu' e si aspetta la pratica."""
    r = rif or {}
    rgn = re.sub(r"[^0-9]", "", str(r.get("rg") or ""))
    scn = _compatta(r.get("sisco") or "")

    # Il ruolo POSSIBILE - il numero di "COMUNICAZIONE n/anno/LAV", che meta'
    # delle volte e' il ruolo - vale solo per CERCARE, e solo se non c'e' un
    # ruolo vero e se un fascicolo con quel numero esiste davvero. Non entra
    # mai in un fascicolo nuovo.
    #
    # "Esiste davvero" vuol dire IN REGISTRO, non "ha gia' una cartella".
    # Guardare il solo cartella_rg lasciava orfane le comunicazioni di
    # cancelleria dei fascicoli che la cartella non ce l'hanno ancora: in
    # archivio erano 28, tutte "COMUNICAZIONE n/anno/LAV", identiche alle 46
    # che si agganciavano - l'unica differenza era se la cartella c'era gia'.
    # La cartella la crea crea_cartella_fascicolo poco piu' sotto.
    # Misurato: sulle 52 PEC che portano SIA il ruolo vero SIA il numero di
    # comunicazione, i due numeri coincidono 52 volte su 52. Le guardie che
    # contano restano tutte: il ruolo conteso rifiuta (19 casi), il ruolo che
    # in registro non c'e' resta orfano (35 casi), la sede diversa rifiuta (2).
    # Dopo il reset conta il doppio: a bacino appena caricato le cartelle non
    # esistono ancora, e con il solo cartella_rg nessuna di queste PEC
    # troverebbe il suo fascicolo.
    if not rgn and r.get("rg_forse"):
        forse = re.sub(r"[^0-9]", "", str(r["rg_forse"]))
        if forse and (forse in FASC["cartella_rg"] or forse in FASC["per_rg"]):
            rgn = forse

    # Il ruolo diviso fra due fascicoli non decide piu' da solo: da qui in
    # avanti vale solo come conferma della pratica, mai come aggancio.
    conteso = FASC["rg_ambigui"].get(rgn) if rgn else None

    # 1. IL NUMERO DI PRATICA, che porta il suffisso ed e' unico per fascicolo.
    # Si guarda per PRIMO davvero: nessun controllo sul ruolo puo' precederlo.
    # Il controllo di sede sta piu' sotto, nel ramo del ruolo, perche' e' li'
    # che serve - e perche' la voce che risponde a un ruolo conteso e' una
    # qualsiasi delle due: farle rifiutare una PEC che porta la pratica giusta
    # significava scartare proprio i casi che questa modifica deve risolvere.
    if scn and scn in FASC["cartella_sisco"]:
        c = FASC["cartella_sisco"][scn]
        marca_cartella(c, FASC["per_sisco"].get(scn), r)
        return c, f"SISCO {r.get('sisco') or scn} (indice diretto)", False

    voce_sisco = (FASC["per_sisco"].get(scn)
                  or FASC["per_sisco"].get(scn.rstrip("0123456789"))) if scn else None

    # 2. il ruolo. Solo adesso, e solo se la pratica non ha gia' risposto:
    # la sede e' l'unico controllo che distingue Siracusa da Catania a
    # parita' di ruolo, quindi vale per tutti gli agganci che partono dal ruolo.
    if not voce_sisco and rgn and not _sede_compatibile(FASC["per_rg"].get(rgn), r):
        v = FASC["per_rg"].get(rgn) or {}
        return None, (f"ruolo {r.get('rg')} trovato ma in sede diversa: la PEC viene da "
                      f"{r.get('sede')}, il fascicolo di {v.get('ricorrente') or ''} "
                      f"e' di {v.get('sede')}"), False

    if not voce_sisco and rgn and not conteso and rgn in FASC["cartella_rg"]:
        c = FASC["cartella_rg"][rgn]
        marca_cartella(c, FASC["per_rg"].get(rgn), r)
        # Il ruolo puo' venire da rg_forse: in QUEL ramo r["rg"] non esiste, ed
        # e' esattamente il caso in cui si entra qui. La parentesi quadra
        # buttava via un fascicolo GIA' TROVATO mentre scriveva come lo aveva
        # trovato: novanta messaggi a ogni giro finivano fra gli illeggibili.
        return c, f"ruolo {r.get('rg') or rgn} (indice diretto)", False

    voce = voce_sisco
    if not voce and rgn and not conteso:
        voce = FASC["per_rg"].get(rgn)
    if not voce and conteso:
        # due rifiuti diversi, perche' chiedono due cose diverse a chi legge:
        # o la pratica manca davvero, o c'era ma in archivio non risponde.
        chi = ', '.join(conteso[:3])
        return None, ((f"ruolo {r.get('rg')} presente su piu' fascicoli ({chi}): "
                       f"la pratica {r.get('sisco')} non risulta in archivio, "
                       f"quindi non si puo' dire quale") if scn else
                      (f"ruolo {r.get('rg')} presente su piu' fascicoli ({chi}): "
                       f"senza il numero di pratica non si puo' dire quale")), False
    if voce and not _sede_compatibile(voce, r):
        da = (f"pratica {r.get('sisco')}" if voce is voce_sisco
              else f"ruolo {r.get('rg')}")
        return None, (f"{da}: fascicolo trovato ma in sede diversa. "
                      f"La PEC viene da {r.get('sede')}, il fascicolo "
                      f"{voce.get('ricorrente') or ''} e' di "
                      f"{voce.get('sede')}"), False
    if voce:
        c = cartella_del_fascicolo(voce)
        if c:
            marca_cartella(c, voce, r)
            return c, f"{r.get('sisco') or r.get('rg')} -> {voce['ricorrente']}", False
        c, nota = crea_cartella_fascicolo(voce)
        if c:
            marca_cartella(c, voce, r)
            return c, f"{r.get('sisco') or r.get('rg')} -> {voce['ricorrente']} ({nota})", True

    if ricorrente and udienza:
        k = _chiave_fascicolo(ricorrente, f"{udienza:%d%m%y}")
        if k and k in FASC["per_chiave"]:
            return FASC["per_chiave"][k], f"{ricorrente} udienza {udienza:%d/%m/%Y}", False
    return None, None, False

# --------------------------------------------------------------------------
# FILE .msg (messaggi Outlook incorporati)
# --------------------------------------------------------------------------
def estrai_da_msg(dati_o_percorso):
    """Un .msg e' un contenitore OLE, non un .eml. Outlook incorpora le PEC
    cosi': i byte grezzi non esistono, l'allegato va salvato e poi aperto.
    Restituisce [(nome, byte), ...] e il testo del corpo."""
    try:
        import olefile
    except ImportError:
        return [], "", "manca la libreria olefile: esegui  pip install olefile"

    try:
        ole = olefile.OleFileIO(dati_o_percorso)
    except Exception as e:
        return [], "", f"non e' un file .msg leggibile: {e}"

    def leggi(percorso):
        try:
            with ole.openstream(percorso) as s:
                return s.read()
        except Exception:
            return None

    def testo_di(prefisso, dentro=()):
        for suffisso, dec in (("001F", "utf-16-le"), ("001E", "cp1252")):
            d = leggi(list(dentro) + [f"__substg1.0_{prefisso}{suffisso}"])
            if d:
                try:
                    return d.decode(dec, "ignore").rstrip("\x00")
                except Exception:
                    pass
        return ""

    fuori = []
    corpo = testo_di("1000")

    def scendi(dentro, prof=0):
        if prof > 4:
            return
        storages = set()
        for voce in ole.listdir(streams=True, storages=True):
            if len(voce) == len(dentro) + 1 and voce[:len(dentro)] == list(dentro) \
               and voce[-1].startswith("__attach_version1.0"):
                storages.add(voce[-1])
        for st in sorted(storages):
            base = list(dentro) + [st]
            nome = testo_di("3707", base) or testo_di("3704", base) or "allegato"
            dati = leggi(base + ["__substg1.0_37010102"])
            if dati:
                fuori.append((nome, dati))
                continue
            # messaggio incorporato dentro l'allegato: si scende ancora
            annidato = base + ["__substg1.0_3701000D"]
            if any(v[:len(annidato)] == annidato for v in ole.listdir(storages=True)):
                scendi(annidato, prof + 1)

    try:
        scendi([])
    finally:
        ole.close()
    return fuori, corpo, None

# --------------------------------------------------------------------------
# OUTLOOK
# --------------------------------------------------------------------------
PR_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
PR_MSGID   = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
PR_ATTDATA = "http://schemas.microsoft.com/mapi/proptag/0x37010102"

def prop(oggetto, nome, dflt=""):
    """Outlook solleva eccezioni su proprieta' che pure esistono: elementi non
    ancora scaricati, ricevute, voci non-mail. Ogni accesso va protetto."""
    try:
        v = getattr(oggetto, nome)
        return dflt if v is None else v
    except Exception:
        return dflt

def e_un_messaggio(m):
    """43 = messaggio di posta, 46 = ricevuta/rapporto di consegna."""
    c = prop(m, "Class", 0)
    return c in (43, 46)

def e_la_busta(nome):
    """Outlook rinomina l'allegato: 'postacert.eml (242 MB).msg'.
    Il confronto esatto non lo riconosceva."""
    n = testo(nome).lower()
    return n.startswith("postacert") or ("postacert" in n and n.endswith((".eml", ".msg")))

def contenuto_allegato(a, cartella_tmp):
    """Tre strade, in ordine: i byte grezzi; il salvataggio su file; e se
    quello che esce e' un .msg, l'apertura del contenitore OLE.
    Restituisce (elenco_file, testo, nota)."""
    nome = pulisci_nome(prop(a, "FileName")) or "allegato"
    try:
        d = a.PropertyAccessor.GetProperty(PR_ATTDATA)
        if d:
            return [(nome, bytes(d))], "", None
    except Exception:
        pass
    # messaggio incorporato: i byte non esistono, va salvato
    try:
        cartella_tmp.mkdir(parents=True, exist_ok=True)
        p = cartella_tmp / pulisci_nome(nome, 80)
        a.SaveAsFile(str(p))
        if not p.exists() or not p.stat().st_size:
            return [], "", f"{nome}: salvataggio vuoto"
        if p.suffix.lower() == ".msg":
            dentro, corpo, err_msg = estrai_da_msg(str(p))
            p.unlink(missing_ok=True)
            if err_msg:
                return [], "", f"{nome}: {err_msg}"
            return dentro, corpo, None
        d = p.read_bytes()
        p.unlink(missing_ok=True)
        return [(nome, d)], "", None
    except Exception as e:
        return [], "", f"{nome}: {e}"

# Se Outlook e' stato avviato da noi lo chiudiamo alla fine: senza, ogni
# lancio ne lascia uno appeso e dopo qualche giro la macchina rallenta.
AVVIATO_DA_NOI = {"si": False, "app": None}

import atexit

def _chiudi_a_fine_programma():
    try:
        chiudi_outlook()
    except Exception:
        pass

atexit.register(_chiudi_a_fine_programma)

def collega_outlook():
    """Outlook NON deve essere aperto: Dispatch lo avvia da solo, senza
    mostrare finestre. Il messaggio di prima diceva il contrario ed e' quello
    che ci ha fatto credere di doverlo tenere aperto.

    Restano due condizioni vere, e diverse fra loro:
      - il profilo non deve chiedere la password, se no la chiamata resta
        appesa ad aspettare una finestra che nessuno vede;
      - in modalita' non memorizzata serve la rete verso Exchange, perche'
        i messaggi non sono sul disco.
    """
    try:
        import win32com.client
    except Exception as e:
        err(f"win32com non disponibile: {e}")
        return None

    # era gia' aperto? GetActiveObject fallisce se non c'e' nessuna istanza
    gia_aperto = True
    try:
        win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        gia_aperto = False

    try:
        app = win32com.client.Dispatch("Outlook.Application")
        ns = app.GetNamespace("MAPI")
        # Logon silenzioso: niente finestra di scelta profilo, niente password.
        # Se il profilo ne chiede una, qui si ferma invece di appendersi.
        try:
            ns.Logon("", "", False, False)
        except Exception:
            pass          # gia' connesso: Logon su una sessione viva solleva
        AVVIATO_DA_NOI["si"] = not gia_aperto
        AVVIATO_DA_NOI["app"] = app
        log("Outlook: " + ("gia' aperto, mi collego"
                           if gia_aperto else "non era aperto, l'ho avviato io"))
        # Appena avviato, Outlook non e' ancora collegato a Exchange: la prima
        # lettura della casella fallisce con "problemi della rete" e il giro
        # si perdeva intero. Si prova la casella e si aspetta, fino a 3 minuti.
        if not outlook_pronto(ns, 180 if not gia_aperto else 60):
            err("Outlook non riesce a collegarsi a Exchange: rete o VPN giu'. "
                "Riprovo al prossimo giro.")
            return None
        return ns
    except Exception as e:
        err(f"collegamento a Outlook fallito: {e}")
        err("Non serve tenerlo aperto, ma il profilo non deve chiedere la "
            "password e il programma non deve girare come amministratore.")
        return None

def outlook_pronto(ns, secondi):
    """Vero quando la casella si lascia leggere. Ogni 15 secondi un tentativo."""
    fine = time.time() + secondi
    prova = 0
    while True:
        prova += 1
        try:
            ns.GetDefaultFolder(6).Items.Count     # 6 = Posta in arrivo
            for f in ns.Folders:
                f.Name
            if prova > 1:
                log(f"Outlook: collegato a Exchange al tentativo {prova}")
            return True
        except Exception as e:
            if time.time() >= fine:
                err(f"Outlook: dopo {prova} tentativi ancora non collegato ({e})")
                return False
            if prova == 1:
                log("Outlook: non ancora collegato a Exchange, aspetto...")
            time.sleep(15)

def chiudi_outlook():
    """Si chiude solo se l'abbiamo avviato noi: se lo stava usando l'utente,
    chiuderglielo sotto le mani sarebbe peggio del processo appeso."""
    if not AVVIATO_DA_NOI["si"] or not AVVIATO_DA_NOI["app"]:
        return
    try:
        AVVIATO_DA_NOI["app"].Quit()
        log("Outlook chiuso: l'avevo avviato io")
    except Exception as e:
        err(f"Outlook non si e' chiuso: {e}")
    finally:
        AVVIATO_DA_NOI["si"] = False
        AVVIATO_DA_NOI["app"] = None

def tutte_le_cartelle(radice, prof=0):
    if prof > 8:
        return
    for i in range(1, radice.Folders.Count + 1):
        try:
            c = radice.Folders.Item(i)
        except Exception:
            continue
        yield c
        yield from tutte_le_cartelle(c, prof + 1)

def filtro_date(cfg):
    """Da data a data invece che 'ultimi N giorni': serve per lavorare un
    periodo alla volta e capire i risultati, invece di un lancio massivo
    indistinto in cui non si distingue piu' nulla."""
    dal, al = cfg.get("dal"), cfg.get("al")
    if dal or al:
        pezzi = []
        if dal:
            d = datetime.strptime(str(dal)[:10], "%Y-%m-%d")
            pezzi.append("[ReceivedTime] >= '" + d.strftime("%d/%m/%Y") + " 00:00'")
        if al:
            a = datetime.strptime(str(al)[:10], "%Y-%m-%d")
            pezzi.append("[ReceivedTime] <= '" + a.strftime("%d/%m/%Y") + " 23:59'")
        return " AND ".join(pezzi), f"dal {dal or 'inizio'} al {al or 'oggi'}"
    g = int(cfg.get("giorni_indietro") or 30)
    d = datetime.now() - timedelta(days=g)
    return ("[ReceivedTime] >= '" + d.strftime("%d/%m/%Y") + " 00:00'",
            f"ultimi {g} giorni (dal {d.strftime('%d/%m/%Y')})")

def scegli_cartelle(ns, nomi, escludi=None):
    escludi = [e.lower() for e in (escludi or [])]
    trovate = []
    for i in range(1, ns.Folders.Count + 1):
        try:
            store = ns.Folders.Item(i)
        except Exception:
            continue
        for c in tutte_le_cartelle(store):
            percorso = testo(prop(c, "FolderPath")).lower().replace("\\", "/")
            if any(e in percorso for e in escludi):
                continue
            if not nomi:
                trovate.append(c)
            elif any(n.lower().strip("/\\").replace("\\", "/") in percorso for n in nomi):
                trovate.append(c)
    return trovate

def chiave_stabile(m):
    """EntryID cambia se la mail viene spostata: uso il Message-ID della PEC,
    che e' assegnato dal gestore e non cambia mai."""
    # NON chiamare "prop" la variabile del ciclo: c'e' gia' una funzione con
    # quel nome, e sovrascriverla fa fallire tutte le righe successive
    for chiave_mapi in (PR_MSGID, PR_HEADERS):
        try:
            v = testo(m.PropertyAccessor.GetProperty(chiave_mapi))
            if chiave_mapi == PR_HEADERS:
                mm = re.search(r"(?im)^Message-ID:\s*(<[^>]+>)", v)
                v = mm.group(1) if mm else ""
            if v.strip():
                return "mid:" + v.strip()
        except Exception:
            pass
    base = "|".join([testo(prop(m, "Subject")),
                     mittente_vero(m),
                     testo(prop(m, "ReceivedTime"))])
    return "hash:" + hashlib.sha256(base.encode("utf-8", "ignore")).hexdigest()[:24]

PREFISSI = [
    r"^\s*(?:[IRF]|fwd?|re|inoltro|inoltrato|trasmissione)\s*:\s*",   # I: R: Fwd: Inoltro:
    r"^\s*istituzionale\s*:\s*",
    r"^\s*urgente\.?\s*:?\s*",
    r"^\s*PEC\s*:\s*",
    r"^\s*notifica di assegnazione pec\s*:\s*",
    r"^\s*notifica di comunicazione[^:]*:\s*",
    r"^\s*posta certificata\s*:\s*",
    r"^\s*NT\s*\[\d+\]\s*",
]

def oggetto_reale(ogg):
    """Toglie gli involucri dell'oggetto finche' resta quello scritto dall'avvocato.
    'I: Notifica di assegnazione PEC: POSTA CERTIFICATA: Notificazione ai sensi...'
    diventa 'Notificazione ai sensi...'."""
    s = testo(ogg).strip()
    cambiato = True
    giri = 0
    while cambiato and giri < 8:
        cambiato = False
        giri += 1
        for p in PREFISSI:
            nuovo = re.sub(p, "", s, count=1, flags=re.I)
            if nuovo != s:
                s = nuovo.strip()
                cambiato = True
    return s

RE_COMUNICAZIONE = re.compile(r"\bcomunicazione\s+n?\.?\s*(\d{1,5})\s*/\s*(\d{4})\s*/\s*LAV\b", re.I)
# la sigla PRIMA del numero ha la precedenza: in "sentenza nr. 270/26 RG 3074"
# il ruolo e' 3074, non 270/26 (che e' il numero di sentenza)
# Il separatore fra numero e anno puo' essere anche uno SPAZIO: "RG001009 24",
# "RG3118 2024". Dopo la sigla R.G. lo spazio non e' ambiguo - nessun'altra
# cosa segue quella sigla - mentre altrove lo sarebbe, e infatti nel ruolo
# scritto DOPO il numero (RE_RUOLO_POST) lo spazio non si accetta.
RE_RUOLO_PRE = re.compile(
    r"R\.?\s?G\.?(?:\s?L\.?)?\.?\s*(?:n\.?|nr\.?)?\s*"
    r"(?:[A-Za-z]{2,12}\s+){0,2}(\d{1,6})\s*[/\-\s]\s*(\d{2,4})\b", re.I)
RE_RUOLO_POST = re.compile(r"(\d{1,6})\s*[/\-]\s*(\d{2,4})\s*R\.?\s?G\.?", re.I)
# l'anno del SISCO e' recente: "222/84/AI" non e' un codice, e' un articolo di legge
RE_SISCO = re.compile(r"\b(\d{1,5})\s*[/\-]\s*(1\d|2\d|3[0-5])\s*[/\-]?\s*(CO|AO|AI)\s*[/\-]?\s*(\d)?\b", re.I)
RE_SPESE = re.compile(r"\b(\d{1,4})\s*[/\-]\s*(\d{2})\s*[/\-]\s*E\b", re.I)
RE_PRATICA = re.compile(r"pratica\s*:?\s*(\d{1,5})\s*[/\-]\s*(\d{2})\s*[/\-]\s*(co|ao|ai)\s*[/\-]?\s*(\d)?", re.I)
RE_SENT = re.compile(r"\bsent(?:enz[ae])?\.?\s*(?:n\.?|nr\.?)?\s*(\d{1,5})\s*[/\-]\s*(\d{2,4})", re.I)

def fase_ctu(ogg_reale, allegati, fasi):
    """Le COMUNICAZIONE n/anno/LAV sono le fasi di una consulenza tecnica:
    ognuna e' un evento del fascicolo, non un fascicolo nuovo."""
    s = (testo(ogg_reale) + " " + " ".join(testo(a) for a in (allegati or []))).lower()
    for fase, parole in (fasi or {}).items():
        if fase.startswith("_"):
            continue
        for p in parole:
            if p and p.lower() in s:
                return fase
    return None

def riferimenti(ogg_reale):
    """Numeri che agganciano la comunicazione a un fascicolo gia' aperto."""
    r = {}
    m = RE_COMUNICAZIONE.search(testo(ogg_reale))
    if m:
        r["comunicazione"] = f"{m.group(1)}/{m.group(2)}/LAV"
        # Misurato: su 1.284 oggetti "COMUNICAZIONE n/anno/LAV", nel 49% dei
        # casi quel numero e' il numero di RUOLO di un fascicolo che esiste.
        # Non basta a creare un fascicolo - nell'altro 51% non corrisponde a
        # niente - ma basta a CERCARLO. Qui diventa un ruolo possibile: se un
        # fascicolo con quel ruolo c'e', la PEC ci si aggancia; se non c'e',
        # non succede nulla e nessun fascicolo nasce con un ruolo inventato.
        r["rg_forse"] = f"{int(m.group(1))}/{m.group(2)}"
    m = RE_RUOLO_PRE.search(testo(ogg_reale)) or RE_RUOLO_POST.search(testo(ogg_reale))
    if m:
        num, anno = m.group(1), m.group(2)
        if len(anno) == 2:
            anno = "20" + anno
        rg = f"{int(num)}/{anno}"
        # IL NUMERO DI COMUNICAZIONE, misurato: su 1.284 oggetti
        # "COMUNICAZIONE n/anno/LAV", nel 49% dei casi quel numero E' il
        # numero di ruolo di un fascicolo che esiste. Nell'altro 51% non
        # corrisponde a niente.
        #
        # Quindi non si decide qui. Mercoledi' l'avevo escluso del tutto e
        # buttavo via 640 agganci buoni; prenderlo sempre per ruolo creerebbe
        # 644 fascicoli con un ruolo inventato. Diventa invece un ruolo
        # POSSIBILE: serve a CERCARE il fascicolo, e non viene mai scritto
        # dentro un fascicolo nuovo. Se un fascicolo con quel ruolo esiste, si
        # aggancia; se non esiste, non succede niente.
        if r.get("comunicazione", "").startswith(f"{int(num)}/{anno}/"):
            r["rg_forse"] = rg
            rg = None
        elif re.search(r"sentenz\w*\s+n\.?\s*0*" + re.escape(num) +
                       r"\s*/\s*" + re.escape(m.group(2)), testo(ogg_reale), re.I):
            # il numero di sentenza: stessa logica, serve a cercare non a creare
            r["rg_forse"] = rg
            rg = None
        if rg:
            r["rg"] = rg
    m = RE_PRATICA.search(testo(ogg_reale)) or RE_SISCO.search(testo(ogg_reale))
    if m:
        sigla = (m.group(3) or "").upper()
        if sigla:
            r["sisco"] = f"{int(m.group(1)):04d}/{m.group(2)}/{sigla}" + (("/" + m.group(4)) if m.group(4) else "")
    m = RE_SPESE.search(testo(ogg_reale))
    if m:
        r["spese_legali"] = f"{int(m.group(1))}/{m.group(2)}/E"
    m = RE_SENT.search(testo(ogg_reale))
    if m:
        r["sentenza"] = f"{int(m.group(1))}/{m.group(2)}"
    return r

def tipo_di_pec(ogg_reale, allegati, tipi):
    """Che genere di comunicazione e': prima istanza, variazione, esito, pagamento, CTU."""
    s = testo(ogg_reale).lower()
    nomi = " ".join(testo(a).lower() for a in (allegati or []))
    punteggi = {}
    for tipo, parole in (tipi or {}).items():
        if tipo.startswith("_"):
            continue
        n_ogg = sum(1 for p in parole if p and p.lower() in s)
        n_all = sum(1 for p in parole if p and p.lower() in nomi)
        if n_ogg or n_all:
            punteggi[tipo] = (n_ogg, n_all)
    if punteggi:
        prio = (tipi.get("_priorita") or {})
        # L'ALLEGATO PESA PIU' DELL'OGGETTO.
        # "Notificazione ai sensi della legge 53" e' la stessa formula per un
        # ricorso e per una sentenza: l'oggetto non distingue. A distinguere e'
        # cio' che c'e' dentro - SENTENZA_CRON_28727930s.pdf - e quello pesa
        # abbastanza da superare la priorita' di prima_istanza, che altrimenti
        # vince sempre (110 contro 60) e fa passare una sentenza per un
        # ricorso nuovo.
        punteggi = {k: o * 10 + a * 120 + prio.get(k, 50)
                    for k, (o, a) in punteggi.items()}
    if not punteggi:
        # l'oggetto porta il numero di ruolo ma non dice che cosa comunica:
        # il genere si stabilisce leggendo corpo e allegati, non l'oggetto
        return "da_leggere" if riferimenti(ogg_reale).get("rg") else "da_classificare"
    return max(punteggi.items(), key=lambda x: x[1])[0]

DOCUMENTALI = (".pdf", ".p7m", ".eml", ".zip", ".doc", ".docx", ".rtf", ".tif", ".tiff")

def allegati_documentali(m):
    """Le firme in calce (image001.png) non contano: cerco veri documenti."""
    try:
        for j in range(1, m.Attachments.Count + 1):
            n = testo(prop(m.Attachments.Item(j), "FileName")).lower()
            if n.endswith(DOCUMENTALI) and not re.match(r"^image\d+\.", n):
                return n
    except Exception:
        pass
    return None

def destinatari_visibili(m):
    """A/CC dell'elemento Outlook. Sugli avvisi del Portale qui compare
    l'utente interno, non la casella PEC: quella si cerca nel corpo."""
    return " ".join([testo(prop(m, "To")), testo(prop(m, "CC"))]).lower()

def e_casella_interna(indirizzo, caselle):
    """Un indirizzo interno all'Istituto non e' il destinatario di una notifica:
    e' chi ha ricevuto un inoltro. Riconoscerlo evita di scambiare l'inoltro
    per l'atto originale."""
    a = testo(indirizzo).lower()
    if not a:
        return False
    for p in caselle.get("interne", ["@inps.it"]):
        if p and p.lower() in a:
            # le caselle istituzionali di ricezione restano tali
            if any(a == c.lower() for c in
                   (caselle.get("valida", []) + caselle.get("impropria", []))):
                return False
            return True
    return bool(re.match(r"^avv\.", a))      # PEC personale di un avvocato dell'ufficio

def casella_di_notifica(m, testo_esteso, caselle, daticert=None, tutti=None):
    """A quale casella l'atto e' stato notificato.

    Ordine: fra tutte le attestazioni si sceglie quella con un destinatario
    istituzionale; se ci sono solo indirizzi interni, la PEC e' un inoltro e
    il destinatario originale non e' ricostruibile - meglio dirlo che
    spacciare l'inoltro per la notifica."""
    elenco = list(tutti or ([daticert] if daticert else []))

    def destinatari(d):
        return (d or {}).get("destinatari_certificati") or (d or {}).get("destinatari") or []

    # 1) un'attestazione che nomina una casella conosciuta
    for d in elenco:
        for a in destinatari(d):
            for c in caselle.get("valida", []):
                if c and c.lower() == a.lower():
                    return {"casella": c, "notifica_valida": True, "difetto": None,
                            "fonte": "daticert.xml"}
            for c in caselle.get("impropria", []):
                if c and c.lower() == a.lower():
                    return {"casella": c, "notifica_valida": False, "fonte": "daticert.xml",
                            "difetto": caselle.get("difetto_predefinito",
                                                   "casella non deputata agli atti giudiziari")}
    # 2) un'attestazione con un destinatario istituzionale non ancora censito
    for d in sorted(elenco, key=lambda x: -(x or {}).get("profondita", 0)):
        for a in destinatari(d):
            if not e_casella_interna(a, caselle):
                return {"casella": a, "notifica_valida": False, "fonte": "daticert.xml",
                        "difetto": caselle.get("difetto_predefinito",
                                               "casella non deputata agli atti giudiziari")}
    # 3) solo indirizzi interni: e' un inoltro
    if elenco:
        interni = sorted({a for d in elenco for a in destinatari(d)})
        return {"casella": None, "notifica_valida": None, "fonte": "daticert.xml",
                "difetto": "PEC inoltrata internamente (" + ", ".join(interni[:2]) +
                           "): la casella di notifica originale non e' ricostruibile"}

    dove = (destinatari_visibili(m) + " " + testo(testo_esteso)).lower()
    for c in caselle.get("valida", []):
        if c and c.lower() in dove:
            return {"casella": c, "notifica_valida": True, "difetto": None, "fonte": "corpo"}
    for c in caselle.get("impropria", []):
        if c and c.lower() in dove:
            return {"casella": c, "notifica_valida": False, "fonte": "corpo",
                    "difetto": caselle.get("difetto_predefinito",
                                           "casella non deputata agli atti giudiziari")}
    m2 = re.search(r"(?:indirizzato\s+a|destinatari?o?\s*:|addressed\s+to)\s*:?\s*"
                   r"([\w.+-]+@[\w.-]+)", dove, re.I)
    if m2 and not e_casella_interna(m2.group(1), caselle):
        return {"casella": m2.group(1), "notifica_valida": False, "fonte": "corpo",
                "difetto": caselle.get("difetto_predefinito",
                                       "casella non deputata agli atti giudiziari")}
    return {"casella": None, "notifica_valida": None, "fonte": None,
            "difetto": "casella di destinazione non individuata"}

_REGEX_CACHE = {}
def regex_forti(criteri):
    chiave = id(criteri)
    if chiave not in _REGEX_CACHE:
        _REGEX_CACHE[chiave] = [(eti, re.compile(p, re.I))
                                for eti, p in criteri.get("oggetto_forte_regex", [])]
    return _REGEX_CACHE[chiave]

# All'inizio dell'oggetto: "Consegna:", "Recapitato:", "ACCETTAZIONE -",
# "Avvenuta consegna", "POSTA CERTIFICATA: Consegna:"... Il gestore mette
# sempre il participio davanti, e quello e' il segno.
RE_RICEVUTA = re.compile(
    r"^\s*(?:posta\s+certificata\s*[:\-]?\s*)?"
    r"(?:(?:avviso|ricevuta|notifica|conferma)\s+(?:di\s+)?)?"
    r"(?:avvenut[ao]\s+|mancat[ao]\s+|non\s+)?"
    r"(?:accettazion\w*|consegn\w*|recapit\w*|protocollazion\w*|"
    r"lettur\w*|lett[oa]|anomalia(?:\s+messaggio)?|preavviso|errore|"
    r"rilevazione\s+virus)"
    r"\s*(?:[:\-\u2013]|$)", re.I)


def mittente_vero(m):
    """Per i messaggi interni Outlook non restituisce l'indirizzo ma
    l'identificativo di Exchange:

        /O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23...)

    Su quello nessun criterio sul mittente puo' funzionare, e gran parte
    della posta di questo ufficio arriva cosi'. L'indirizzo vero si chiede
    con la proprieta' SMTP, e se non c'e' si ripiega sul nome."""
    mit = testo(prop(m, "SenderEmailAddress"))
    if mit and not mit.startswith("/"):
        return mit
    # PR_SENT_REPRESENTING_SMTP_ADDRESS / PR_SENDER_SMTP_ADDRESS
    for chiave in ("http://schemas.microsoft.com/mapi/proptag/0x5D01001F",
                   "http://schemas.microsoft.com/mapi/proptag/0x5D02001F",
                   "http://schemas.microsoft.com/mapi/proptag/0x0065001F"):
        try:
            v = testo(m.PropertyAccessor.GetProperty(chiave)).strip()
            if v and "@" in v:
                return v
        except Exception:
            pass
    try:
        u = m.Sender.GetExchangeUser()
        if u:
            v = testo(u.PrimarySmtpAddress).strip()
            if v:
                return v
    except Exception:
        pass
    return mit or testo(prop(m, "SenderName"))

def valuta(m, criteri, usa_corpo=True):
    """Decide se la PEC interessa e per quale motivo.

    Segnale FORTE  -> basta da solo.
    Segnale DEBOLE -> ne servono due, oppure uno piu' un mittente PEC.
    Cosi' 'sentenza' dentro una mail interna non fa scattare nulla, e un
    gestore PEC generico come legalmail.it non trascina dentro i patronati."""
    ogg = testo(prop(m, "Subject")).lower()
    reale = oggetto_reale(ogg)
    mit = mittente_vero(m).lower()

    # Il modulo di richiesta istruttoria porta sede, pratica e nominativo in
    # forma rigida: e' la fonte piu' pulita che abbiamo. Scavalca la parola
    # "istruttoria" che lo escluderebbe, ma NON le altre esclusioni: una
    # ricevuta di consegna resta una ricevuta anche se cita una pratica.
    # LE RICEVUTE DEL GESTORE, per FORMA e non per elenco.
    # Elencare le parole una a una e' una rincorsa che si perde sempre: basta
    # che il gestore ne usi una nuova - "Recapitato:" invece di "Consegna:" -
    # e rientra tutto. Qui si riconosce la forma: un participio di consegna
    # all'inizio dell'oggetto, seguito da due punti o dal messaggio originale.
    # Viene PRIMA del modulo: una ricevuta di consegna che cita la pratica
    # resta una ricevuta, non diventa la comunicazione originale.
    if RE_RICEVUTA.match(ogg):
        return False, "ricevuta del gestore (non e' una comunicazione)"

    # Le risposte automatiche. L'esclusione in elenco e' "risposta:", ma
    # l'oggetto vero dice "Risposta automatica: TRASM. RICHIESTA ISTRUTTORIA
    # (Sede: SIRACUSA - Pratica: ...)": fra "risposta" e i due punti c'e' una
    # parola, e il confronto per sottostringa non la prende. Peggio: il modulo
    # citato nell'oggetto la faceva entrare come comunicazione valida.
    # Anche qui la forma, non l'elenco: qualunque cosa cominci con "risposta"
    # seguita da al piu' due parole e dai due punti e' una risposta.
    if re.match(r"^\s*(?:r|risposta|reply|out of office|fuori sede)"
                r"(?:\s+\w+){0,2}\s*:", ogg, re.I) \
       and re.search(r"automatic|assenza|fuori sede|out of office|ferie|"
                     r"non sono in ufficio|rispondo", ogg, re.I):
        return False, "risposta automatica"

    # LA RISPOSTA DELLA UO IN FORMA DI RISPOSTA. Non tutte le istruttorie si
    # chiudono con la trasmissione formale: moltissime la UO le chiude
    # rispondendo alla nostra mail, e allora l'oggetto e' "R: RICHIESTA
    # ISTRUTTORIA (Sede: ... - Pratica: ...)". Il criterio di esclusione
    # "risposta:" - 1.645 messaggi scartati nell'arretrato - se le portava
    # via tutte, e nel registro quelle istruttorie restavano in attesa per
    # sempre.
    #
    # Si riconosce togliendo dall'oggetto i prefissi di risposta e di inoltro:
    # quello che resta deve essere la richiesta di istruttoria con la sua
    # pratica. Le risposte AUTOMATICHE restano fuori: si guardano prima.
    # Il nostro marcatore basta da solo: se l'oggetto porta [RIF pratica] e'
    # una mail che nasce da un nostro sollecito, e le mail ordinarie della UO
    # non hanno niente che le faccia scattare - non sono PEC, non citano il
    # ruolo, spesso dicono solo "vedi sotto".
    if RE_RIF_SOLL.search(ogg):
        return True, "forte: risposta a un nostro sollecito ([RIF])"

    nudo = re.sub(r"^(?:\s*(?:r|re|rif|risposta|i|fw|fwd|inoltro)\s*:)+\s*", "", ogg, flags=re.I)
    if RE_ISTR_RICH.search(nudo) and RE_ISTR_PRAT.search(nudo):
        if nudo != ogg:
            return True, "forte: risposta della UO a una richiesta istruttoria"
        if RE_ISTR_TRASM.search(nudo):
            return True, "forte: trasmissione istruttoria dalla UO"
        return True, "forte: richiesta istruttoria con pratica"

    # LA TRASMISSIONE DELLA UO. Il referto del 20/09 l'ha detto in chiaro:
    # "1 escluso: 'trasm. richiesta istruttoria'". Il criterio di esclusione
    # buttava via proprio la PEC con cui la UO dice di aver lavorato la
    # richiesta - il solo segnale che chiude l'istruttoria.
    #
    # L'eccezione per il modulo non bastava: RE_MODULO pretende sede, pratica
    # E nominativo fra parentesi, ma Outlook taglia gli oggetti lunghi
    # ("...Pratica: 973/25/CO/1 - ") e senza il nominativo il modulo non si
    # riconosce piu'. Qui si pretende solo cio' che serve: la formula della
    # trasmissione e un numero di pratica. Le ricevute del gestore e le
    # risposte automatiche restano fuori, perche' si guardano prima di questa.
    if RE_ISTR_TRASM.search(ogg) and RE_ISTR_PRAT.search(ogg):
        return True, "forte: trasmissione istruttoria dalla UO"

    ha_modulo = bool(RE_MODULO.search(ogg))

    for x in criteri.get("escludi_oggetto", []):
        if not x:
            continue
        if ha_modulo and re.search(r"istruttoria|oggetto aggiunto", x, re.I):
            continue
        if x.lower() in ogg:
            # le ricevute del gestore si contano a parte: sono la maggior parte
            # del volume e vederle mescolate agli altri scarti confonde
            if re.search(r"accettazion|consegna|lettura|ricevuta|anomalia|"
                         r"non recapit|undeliverable|delivery", x, re.I):
                return False, "ricevuta del gestore (non e' una comunicazione)"
            return False, f"escluso: '{x}'"

    if ha_modulo:
        return True, "forte: modulo con sede, pratica e nominativo"

    for x in criteri.get("oggetto_forte", []):
        if not x:
            continue
        # Un segnale corto - "r.g.", "rg n" - dentro una frase qualsiasi non e'
        # un segnale forte: sono due o tre caratteri che capitano dappertutto,
        # e da soli facevano entrare 333 PEC di sollecito e di pagamento.
        # Sotto i sette caratteri si pretende che sia un ruolo VERO, cioe'
        # seguito da un numero: "R.G. 1250/2026" si', "r.g." in mezzo a una
        # frase no. Il modello del ruolo sta gia' fra le regex forti.
        if len(x) <= 6:
            if re.search(re.escape(x.lower()) + r"\s*\.?\s*\d", ogg):
                return True, f"forte: oggetto '{x}' con il numero"
            continue
        if x.lower() in ogg:
            return True, f"forte: oggetto '{x}'"
    for eti, rx in regex_forti(criteri):
        if rx.search(ogg):
            return True, f"forte: {eti} nell'oggetto"
    for x in criteri.get("mittenti_forti", []):
        if x and x.lower() in mit:
            return True, f"forte: mittente '{x}'"
    dest = destinatari_visibili(m)
    for x in criteri.get("destinatari_forti", []):
        if x and x.lower() in dest:
            return True, "forte: indirizzata alla casella degli atti giudiziari"

    # I colleghi che smistano i ricorsi: la loro posta non e' una notifica, e'
    # l'assegnazione, e porta i fascicoli scritti nell'oggetto. Vale solo se
    # l'oggetto nomina un numero di pratica o un ruolo: senza quel vincolo
    # entrerebbe tutta la corrispondenza d'ufficio.
    # I colleghi che smistano i ricorsi. Con un riferimento nell'oggetto, la
    # loro posta entra come segnale forte.
    #
    # Se il riferimento non c'e' NON si scarta qui: prima si guardano i segnali
    # deboli, come per chiunque altro. Un inoltro e' spesso l'unico modo in cui
    # una comunicazione arriva - l'avvocato la gira al reparto - e uscire dalla
    # valutazione a meta' significava buttare posta vera solo perche' l'aveva
    # rimandata un collega. Chi manda conta per ultimo, dopo che si e' guardato
    # cosa c'e' dentro.
    collega_senza_rif = None
    for x in criteri.get("mittenti_ufficio", []):
        if not x or x.lower() not in mit:
            continue
        if re.search(r"\d{1,4}\s*/\s*\d{2}\s*/\s*(?:CO|AO|CV|AI|IC)|"
                     r"\bR\.?\s*G\.?\s*n?\.?\s*\d{1,6}\s*/\s*\d{2,4}|"
                     r"\b\d{1,6}\s*/\s*(?:19|20)\d{2}\s*/\s*(?:LAV|CIV|PREV)", ogg, re.I):
            return True, f"forte: collega che smista ricorsi ({x.split('@')[0]})"
        collega_senza_rif = x.split("@")[0]
        break

    deboli = []
    for x in criteri.get("oggetto_debole", []):
        if x and x.lower() in reale:
            deboli.append(f"oggetto '{x}'")
    mitt_pec = next((x for x in criteri.get("mittenti_deboli", [])
                     if x and x.lower() in mit), None)
    if mitt_pec:
        deboli.append(f"mittente '{mitt_pec}'")
    dest_pec = next((x for x in criteri.get("destinatari_deboli", [])
                     if x and x.lower() in dest), None)
    if dest_pec:
        deboli.append(f"destinatario '{dest_pec}'")
    # l'allegato rinforza, non sostituisce: serve gia' un indizio nell'oggetto,
    # altrimenti qualunque PEC con un PDF entrerebbe
    if any(d.startswith("oggetto") for d in deboli):
        doc = allegati_documentali(m)
        if doc:
            deboli.append(f"allegato '{doc[:32]}'")
    if usa_corpo and criteri.get("corpo") and len(deboli) == 1:
        # il corpo si legge solo se serve a confermare: e' l'operazione piu' lenta
        corpo = testo(prop(m, "Body"))[:4000].lower()
        for x in criteri["corpo"]:
            if x and x.lower() in corpo:
                deboli.append(f"corpo '{x}'")
                break

    # almeno un indizio deve venire dall'OGGETTO: due segnali puramente tecnici
    # (mittente PEC + destinatario PEC) descrivono il mezzo, non il contenuto
    ha_oggetto = any(d.startswith("oggetto") or d.startswith("corpo") for d in deboli)
    if len(deboli) >= 2 and ha_oggetto:
        return True, "deboli: " + " + ".join(deboli[:2])
    if len(deboli) >= 2:
        return False, "solo segnali tecnici, nulla nell'oggetto"
    if collega_senza_rif:
        return False, ("posta fra colleghi senza riferimenti nell'oggetto")
    if len(deboli) == 1:
        return False, f"un solo segnale debole ({deboli[0]})"
    return False, "nessun criterio soddisfatto"

# --------------------------------------------------------------------------
# I FATTI CHE UNA PEC PORTA
# --------------------------------------------------------------------------
# Una PEC su un fascicolo che esiste non e' posta da archiviare: e' un
# AGGIORNAMENTO di quel fascicolo. Fino a ieri di quell'aggiornamento si
# leggeva soltanto il rinvio, e solo quando stava scritto nel verbale. Tutto
# il resto - la sentenza, il suo esito, le spese, la CTU, i termini concessi,
# il passaggio in appello, un ruolo che prima non c'era - restava dentro i PDF
# allegati, e andava ribattuto a mano due volte: in SISCO e nel registro.
# Parcheggiare la PEC nel fascicolo senza aprire gli allegati non aggiorna
# niente: sposta soltanto il lavoro piu' avanti.
#
# Qui gli allegati si leggono UNA volta e se ne tira fuori un elenco di FATTI,
# lo stesso elenco per ogni PEC, qualunque strada abbia fatto per arrivare.
# Ogni fatto porta con se' quattro cose:
#   tipo       cosa e' (udienza, sentenza, esito, spese, ctu, termine, ...)
#   valore     il dato da scrivere nel fascicolo
#   etichetta  come si legge in italiano, per il pulsante
#   frase      la riga dell'atto da cui viene: senza questa non e' verificabile
#   certezza   'alta' se scritto in forma fissa, 'da_verificare' se dedotto
# Il registro li mostra come pulsanti accanto alla comunicazione: quelli di
# certezza alta li applica da solo, gli altri li chiede.

_GG = r"(\d{1,2})\s*[/.\-]\s*(\d{1,2})\s*[/.\-]\s*(\d{2,4})"
_GG_TESTO = r"(\d{1,2})\s+(" + "|".join(MESI) + r")\s+(\d{4})"

def _frase(t, a, b):
    return re.sub(r"\s+", " ", testo(t)[max(0, a - 45): b + 20]).strip()

def _data_dopo(t, prima, dopo=None):
    """La prima data che segue una delle formule indicate. (data, frase).
    `prima` non deve contenere gruppi catturanti: i tre gruppi sono la data."""
    for coda in (_GG, _GG_TESTO):
        for m in re.finditer(prima + r"[^\n0-9]{0,70}" + coda, testo(t), re.I):
            d = _data_da(m)
            if d and (not dopo or d >= dopo):
                return d, _frase(t, m.start(), m.end())
    return None, None

def _cerca(t, rx):
    m = re.search(rx, testo(t), re.I)
    return (m, _frase(t, m.start(), m.end())) if m else (None, None)

# L'esito, nelle parole del dispositivo. L'ordine conta: "accoglie in parte"
# e' un esito diverso da "accoglie", e va riconosciuto prima.
ESITI = [
    ("parziale",       "accolto in parte",
     r"accogli\w*\s+(?:solo\s+)?(?:in\s+)?part\w|parzialment\w+\s+accogli|accoglimento\s+parziale"),
    ("cessata_materia", "cessata materia del contendere",
     r"cessat\w+\s+(?:la\s+)?materia\s+del\s+contendere"),
    ("estinto",        "estinzione del giudizio",
     r"dichiar\w*\s+(?:l['\u2019]\s*)?estin\w+|\bestinzione\s+del\s+(?:giudizio|processo)"),
    ("cancellata",     "causa cancellata dal ruolo",
     r"cancellazione\s+della\s+causa\s+dal\s+ruolo|cancella\s+la\s+causa\s+dal\s+ruolo"),
    ("inammissibile",  "ricorso inammissibile", r"inammissibil\w+"),
    ("improcedibile",  "ricorso improcedibile", r"improcedibil\w+"),
    ("incompetenza",   "incompetenza",
     r"dichiar\w*\s+(?:la\s+)?(?:propria\s+)?incompetenz\w+"),
    ("conciliato",     "conciliazione / transazione",
     r"verbale\s+di\s+conciliazion|\bconcilia\w*\s+le\s+parti|intervenut\w+\s+transazion"),
    ("rigettato",      "ricorso rigettato",
     r"\brigett\w+\s+(?:il\s+)?ricors|\brespinge\s+il\s+ricors|dichiar\w*\s+infondat\w+\s+il\s+ricors|\brigetta\b"),
    ("accolto",        "ricorso accolto",
     r"\baccogli\w*\s+(?:il\s+)?ricors|in\s+accoglimento\s+del\s+ricors|dichiar\w*\s+il\s+diritto\s+d\w+\s+ricorrent|condann\w+\s+l['\u2019]?\s*I\.?N\.?P\.?S"),
]

RE_SENT_TESTO = re.compile(
    r"\bsentenz\w*\s*(?:n\.?|nr\.?|numero)?\s*(\d{1,5})\s*[/\-]\s*(\d{2,4})", re.I)
RE_SPESE_EURO = re.compile(
    r"spes\w+[^.\n]{0,120}?(?:euro|\u20ac)\s*([\d]{1,3}(?:[.\u00a0]\d{3})*(?:,\d{2})?)"
    r"|(?:liquid\w+|condann\w+)[^.\n]{0,120}?(?:euro|\u20ac)\s*"
    r"([\d]{1,3}(?:[.\u00a0]\d{3})*(?:,\d{2})?)[^.\n]{0,60}?spes", re.I)
RE_CTU_NOME = re.compile(
    r"(?:C\.?T\.?U\.?|consulent\w+\s+tecnic\w+(?:\s+d['\u2019]\s*ufficio)?)[^\n]{0,40}?"
    r"(?:dott|dr|prof|ing|sig)\.?(?:ssa)?\s*"
    r"([A-Z][A-Za-z\u00c0-\u017f'\u2019]{2,}(?:\s+[A-Z][A-Za-z\u00c0-\u017f'\u2019]{2,}){0,2})")
RE_APPELLO = re.compile(
    # "corte d'appello" da sola non basta: compare nell'intestazione di ogni
    # atto di secondo grado e in ogni sentenza che ne citi una. Serve l'ATTO.
    r"atto\s+di\s+appello|ricorso\s+in\s+appello|propon\w+\s+appello|"
    r"appella\s+la\s+sentenz|atto\s+di\s+reclamo|propon\w+\s+reclamo", re.I)
RE_GIUDICE_TESTO = re.compile(
    r"(?:il|la)\s+giudic\w*[^\n]{0,20}?(?:dott|dr)\.?(?:ssa)?\s*"
    r"([A-Z][A-Za-z\u00c0-\u017f'\u2019]{2,}(?:\s+[A-Z][A-Za-z\u00c0-\u017f'\u2019]{2,}){0,2})")

# le descrizioni del modulo che non siamo riusciti a leggere, per forma:
# le stampa il referto, ed e' da qui che si allarga il vocabolario
DESCRIZIONI_MUTE = {}

def _fatto(tipo, valore, etichetta, frase, certezza="alta", **extra):
    v = {"tipo": tipo, "valore": valore, "etichetta": etichetta,
         "frase": testo(frase)[:220] or None, "certezza": certezza}
    v.update(extra)
    return v

# --------------------------------------------------------------------------
# LA DESCRIZIONE DELLA CANCELLERIA: un vocabolario chiuso, non testo libero.
# --------------------------------------------------------------------------
# Il campo "Descrizione:" del modulo dice COSA stanno comunicando, e lo dice
# con frasi fatte: "RINVIATO AD ALTRA UDIENZA DI DISCUSSIONE IL ...",
# "SENTENZA EX ART. 429, I comma CPC NUMERO 135/2026 (Rigetto)",
# "CONSULENZA CTU LEONE MARIO DEPOSITATA". Vale per i rinvii come per le
# sentenze e le CTU.
# Le regex del testo libero, tarate sui PDF, su queste frasi non arrivavano:
# RE_SENT_TESTO vuole il numero subito dopo "sentenza" e qui in mezzo c'e'
# "EX ART. 429, I comma CPC"; la CTU la cercava come nomina ("si dispone
# CTU") e qui e' un deposito gia' avvenuto. Misurato su 10 descrizioni vere:
# 7 rinvii letti, 2 sentenze e 1 CTU non lette.
# Allargare quelle regex avrebbe messo a rischio la lettura dei PDF, dove
# "sentenza" e un numero a quaranta caratteri di distanza non vogliono dire
# niente. Qui il vocabolario e' chiuso, quindi le frasi si riconoscono per
# quello che sono.
RE_MOD_SENT = re.compile(
    r"\b(?:SENTENZA|DECRETO|ORDINANZA)\b[^\n]{0,70}?"
    r"\b(?:NUMERO|NR\.?|N\.)\s*(\d{1,5})\s*/\s*(\d{2,4})", re.I)
# l'esito fra parentesi in coda: "(Rigetto)", "(Accoglimento)", "(Altro)"
RE_MOD_ESITO = re.compile(r"\(\s*(accoglimento[^)]*|rigetto[^)]*|"
                          r"accolt\w+[^)]*|respint\w+[^)]*)\s*\)", re.I)
RE_MOD_CTU_DEP = re.compile(
    r"\b(?:CONSULENZA|RELAZIONE|ELABORATO|PERIZIA)\b[^\n]{0,40}?"
    r"\bC\.?T\.?U\.?\b[^\n]{0,60}?\bDEPOSITAT\w+"
    r"|\bC\.?T\.?U\.?\b[^\n]{0,60}?\bDEPOSITAT\w+", re.I)
RE_MOD_CTU_NOME = re.compile(
    r"\bC\.?T\.?U\.?\s+([A-Z][A-Za-zÀ-ſ'’]{2,}"
    r"(?:\s+[A-Z][A-Za-zÀ-ſ'’]{2,}){0,2})", re.I)

def fatti_dalla_descrizione(descr, ricez=None):
    """I fatti che la descrizione del modulo dice in chiaro. Non sostituisce
    leggi_atto: gli si aggiunge, e i doppioni li scarta chi unisce."""
    t = testo(descr)
    if not t:
        return []
    fatti = []
    m = RE_MOD_SENT.search(t)
    if m:
        num = f"{int(m.group(1))}/{m.group(2)}"
        fatti.append(_fatto("sentenza", num, f"sentenza {num}", t, "alta",
                            data=ricez.isoformat() if hasattr(ricez, "isoformat") else None,
                            data_certa=False))
    e = RE_MOD_ESITO.search(t)
    if e:
        parola = e.group(1).strip().lower()
        if parola.startswith(("rigetto", "respint")):
            fatti.append(_fatto("esito", "rigetto", "ricorso rigettato", t))
        elif parola.startswith(("accoglimento", "accolt")):
            fatti.append(_fatto("esito", "accolto", "ricorso accolto", t))
    if RE_MOD_CTU_DEP.search(t):
        nome = RE_MOD_CTU_NOME.search(t)
        # il deposito e' AVVENUTO: la data e' quella della comunicazione, non
        # un termine da rispettare. Sono due cose diverse e vanno in due fatti
        # diversi, altrimenti finisce negli adempimenti da fare.
        fatti.append(_fatto("ctu_depositata",
                            nome.group(1).strip() if nome else True,
                            "relazione CTU depositata" +
                            (f": {nome.group(1).strip()}" if nome else ""), t,
                            "alta",
                            data=ricez.isoformat() if hasattr(ricez, "isoformat") else None))
    return fatti

# ==========================================================================
# LA PIRAMIDE DELLE LETTURE
# --------------------------------------------------------------------------
# Un solo comando, tanti esecutori. Il comando (leggi_atto) tiene il dossier
# e decide l'ordine; ogni lettore riceve SOLO quello che gli serve, risponde
# e non scrive niente. Chi scrive nell'archivio e' il registro, che legge il
# dossier: qui non si tocca nessun fascicolo.
#
# Gli ordini sono sequenziali perche' dipendono l'uno dall'altro:
#   fase 1  CHI E'      - pratica, ruolo, giudice
#   fase 2  CHE ATTO E' - sentenza, rinvio, CTU, trasmissione della UO
#   fase 3  COSA DICE   - interrogati i lettori dell'atto riconosciuto
#   fase 4  COSA SEGUE  - controlli di coerenza sulle date trovate
#
# Ogni lettore porta la sua targa: nel dossier resta scritto chi ha prodotto
# ogni dato. Quando un dato e' sbagliato si sa quale lettore correggere, e si
# corregge quello solo.
#
# E un rifiuto e' un rifiuto SCRITTO: "numero di sentenza non trovato" non e'
# la stessa cosa di una lettura mai tentata, e un campo vuoto non distingue
# le due. Meta' degli errori di questa settimana venivano da questo.
LETTORI = []

def lettore(nome, fase, targa):
    def dentro(fn):
        fn.nome, fn.fase, fn.targa = nome, fase, targa
        LETTORI.append(fn)
        return fn
    return dentro

def _timbra(fatti, fn):
    for f in (fatti or []):
        f["lettore"] = fn.nome
        f["targa"] = fn.targa
    return fatti or []

ULTIMO_DOSSIER = {}
# Quanto ha prodotto ogni lettore, e cosa ha cercato senza trovare. Serve al
# referto: un lettore si giudica su questi due numeri, non sull'impressione.
CONTA_LETTORI = {}
CONTA_RIFIUTI = {}
CONTA_ATTI = {}
LETTURE = {"scritte": 0}

@lettore("identita", 1, "2026-09-20")
def l_identita(d):
    """Chi e': il numero di pratica e il ruolo, ovunque stiano scritti.
    Viene per primo perche' tutti gli altri lo ricevono e non lo ricercano."""
    fatti, rif = [], d["rif"]
    r2 = riferimenti(d["tutto"][:30000]) or {}
    for chiave, eti in (("rg", "numero di ruolo"), ("sisco", "codice SISCO"),
                        ("spese_legali", "pratica spese legali")):
        if r2.get(chiave) and not rif.get(chiave):
            fatti.append(_fatto("riferimento", r2[chiave],
                                f"{eti} {r2[chiave]}", None, "da_verificare",
                                campo=chiave))
    mrif = RE_RIF_SOLL.search(testo(d.get("oggetto")))
    if mrif:
        prat = re.sub(r"\s+", "", mrif.group(1))
        d["rif_sollecito"] = prat
        if not rif.get("sisco"):
            fatti.append(_fatto("riferimento", prat,
                                f"pratica dal marcatore del sollecito {prat}",
                                testo(d.get("oggetto"))[:200], "alta", campo="sisco"))
    d["pratica"] = rif.get("sisco") or (mrif and re.sub(r"\s+", "", mrif.group(1))) or r2.get("sisco")
    d["ruolo"] = rif.get("rg") or r2.get("rg")
    if not d["pratica"]:
        d["rifiuti"].append({"lettore": "identita", "cercato": "numero di pratica",
                             "esito": "non trovato"})
    gd = d["trib"].get("giudice") or d["xmlatto"].get("giudice")
    if not gd:
        m = RE_GIUDICE_TESTO.search(d["tutto"])
        if m:
            gd = m.group(1).strip()
            fatti.append(_fatto("giudice", gd, f"giudice: {gd}",
                                _frase(d["tutto"], m.start(), m.end()), "da_verificare"))
    d["giudice"] = gd
    return fatti

RE_ISTR_TRASM = re.compile(r"TRASM\.?\s*RICHIESTA\s+ISTRUTTORIA", re.I)
RE_ISTR_RICH = re.compile(r"RICHIESTA\s+ISTRUTTORIA", re.I)
# Il marcatore che il registro mette nell'oggetto del sollecito:
# "... [RIF 661/26/CO/1]". Torna indietro con la risposta della UO, perche'
# l'oggetto sopravvive alle risposte mentre il corpo viene tagliato. E' il
# riferimento piu' affidabile che abbiamo: l'ha scritto il registro stesso.
RE_RIF_SOLL = re.compile(r"\[\s*RIF\s+(\d{1,5}\s*/\s*\d{2}\s*/\s*"
                         r"(?:CO|AO|CV)(?:\s*/\s*\d)?)\s*\]", re.I)
RE_ISTR_RISP = re.compile(r"^\s*(?:r|re|rif|risposta)\s*:", re.I)
RE_ISTR_PRAT = re.compile(r"pratica\s*:\s*(\d{1,5}\s*/\s*\d{2}\s*/\s*"
                          r"(?:CO|AO|CV)(?:\s*/\s*\d)?)", re.I)

@lettore("atto", 2, "2026-09-20")
def l_tipo_atto(d):
    """Che atto e'. Non produce fatti: decide a chi tocca parlare."""
    o = testo(d.get("oggetto"))
    t = d["tutto"]
    if RE_ISTR_TRASM.search(o):
        d["atto"] = "trasmissione_uo"
    elif RE_RIF_SOLL.search(o) and RE_ISTR_RISP.search(o):
        # il nostro sollecito tornato indietro con Rispondi: e' la UO
        d["atto"] = "risposta_uo"
    elif RE_ISTR_RICH.search(o) and RE_ISTR_RISP.search(o):
        d["atto"] = "risposta_uo"
    elif RE_ISTR_RICH.search(o):
        # la richiesta nuda: potrebbe essere la nostra, girata indietro. Dice
        # che dell'istruttoria si sta parlando, non che e' stata definita.
        d["atto"] = "istruttoria_citata"
    elif d["rif"].get("sentenza") or RE_SENT_TESTO.search(t):
        d["atto"] = "sentenza"
    elif re.search(r"C\.?T\.?U", t, re.I):
        d["atto"] = "ctu"
    elif re.search(r"rinvi\w+|differit\w+", t, re.I):
        d["atto"] = "rinvio"
    else:
        d["atto"] = "ignoto"
        d["rifiuti"].append({"lettore": "atto", "cercato": "tipo di atto",
                             "esito": "non riconosciuto"})
    return []

@lettore("istruttoria", 3, "2026-09-20")
def l_istruttoria(d):
    """La risposta della UO, nelle due forme in cui arriva: la trasmissione
    formale ("TRASM. RICHIESTA ISTRUTTORIA") e la risposta alla nostra mail
    ("R: RICHIESTA ISTRUTTORIA"). Sono entrambe la UO che dice di aver
    lavorato la richiesta, e la seconda e' la piu' frequente.

    La richiesta NUDA invece non chiude niente: potrebbe essere la nostra,
    tornata indietro in un inoltro. Si riporta come da verificare, e a
    decidere e' chi guarda."""
    if d["atto"] not in ("trasmissione_uo", "risposta_uo", "istruttoria_citata"):
        return []
    o = testo(d.get("oggetto"))
    m = RE_RIF_SOLL.search(o) or RE_ISTR_PRAT.search(o)
    prat = re.sub(r"\s+", "", m.group(1)) if m else d.get("pratica")
    if not prat:
        d["rifiuti"].append({"lettore": "istruttoria",
                             "cercato": "pratica nella richiesta",
                             "esito": "oggetto riconosciuto ma pratica illeggibile"})
        return []
    if d["atto"] == "istruttoria_citata":
        d["rifiuti"].append({"lettore": "istruttoria", "cercato": "chi parla",
                             "esito": "richiesta citata, non una risposta: da verificare"})
        return [_fatto("istruttoria_definita", prat,
                       f"istruttoria citata (pratica {prat})", o[:200], "da_verificare")]
    come = "trasmessa" if d["atto"] == "trasmissione_uo" else "per risposta"
    return [_fatto("istruttoria_definita", prat,
                   f"istruttoria definita dalla UO, {come} (pratica {prat})",
                   o[:200], "alta")]

# Le PEC della UO sull'istruttoria. Il parere cita spesso sentenze vecchie
# ("sentenza n. 354/2014") e parole del dispositivo: lette come sentenza,
# chiudevano fascicoli con udienza nel 2027. Da queste si prende solo
# l'istruttoria.
ATTI_UO = ("trasmissione_uo", "risposta_uo", "istruttoria_citata")
FATTI_DI_SENTENZA = ("sentenza", "esito", "spese")

@lettore("sentenza", 3, "2026-09-27")
def l_sentenza(d):
    """Numero, data di pubblicazione, esito e spese: sono lo stesso documento,
    quindi un solo lettore. La data scritta nell'atto batte la data della PEC."""
    if d.get("atto") in ATTI_UO:
        return []
    tutto, rif, ricez = d["tutto"], d["rif"], d["ricez"]
    fatti = []
    num, frase_s = rif.get("sentenza"), None
    if not num:
        m = RE_SENT_TESTO.search(tutto)
        if m:
            num = f"{int(m.group(1))}/{m.group(2)}"
            frase_s = _frase(tutto, m.start(), m.end())
    if num:
        pub, frase_p = _data_dopo(tutto, r"(?:pubblicat\w+|depositat\w+)\s+(?:il|in\s+data)")
        fatti.append(_fatto("sentenza", num, f"sentenza {num}", frase_s or frase_p,
                            "alta" if frase_s or rif.get("sentenza") else "da_verificare",
                            data=(pub or ricez).isoformat() if (pub or ricez) else None,
                            data_certa=bool(pub)))
        if not pub:
            d["rifiuti"].append({"lettore": "sentenza",
                                 "cercato": "data di pubblicazione",
                                 "esito": "non scritta: vale la data della PEC"})
    elif d["atto"] == "sentenza":
        d["rifiuti"].append({"lettore": "sentenza", "cercato": "numero di sentenza",
                             "esito": "non trovato"})
    esito = None
    for chiave, etichetta, rx in ESITI:
        m, frase_e = _cerca(tutto, rx)
        if m:
            esito = chiave
            fatti.append(_fatto("esito", chiave, etichetta, frase_e,
                                "da_verificare" if not num else "alta"))
            break
    if num and not esito:
        d["rifiuti"].append({"lettore": "sentenza", "cercato": "esito",
                             "esito": "dispositivo non riconosciuto"})
    m, frase_sp = _cerca(tutto, r"spes\w+[^.\n]{0,40}?(?:integralment\w+\s+)?compensat|"
                                r"compens\w+\s+(?:integralmente\s+)?le\s+spes")
    if m:
        fatti.append(_fatto("spese", "compensate", "spese compensate", frase_sp))
    else:
        m, frase_sp = _cerca(tutto, r"spes\w+[^.\n]{0,30}?irripetibil")
        if m:
            fatti.append(_fatto("spese", "irripetibili", "spese irripetibili", frase_sp))
        else:
            m = RE_SPESE_EURO.search(tutto)
            if m:
                imp = m.group(1) or m.group(2)
                fatti.append(_fatto("spese", imp, f"spese liquidate \u20ac {imp}",
                                    _frase(tutto, m.start(), m.end()), "da_verificare"))
    d["sentenza"] = num
    return fatti

@lettore("coerenza", 4, "2026-09-20")
def l_coerenza(d):
    """Cosa segue dalle date trovate. Non aggiunge fatti: contesta quelli che
    non stanno in piedi, cosi' il registro sa di non doverli applicare."""
    ricez = d["ricez"]
    for f in d["fatti"]:
        if f.get("tipo") == "udienza" and ricez and testo(f.get("valore")) < str(ricez):
            f["certezza"] = "da_verificare"
            d["rifiuti"].append({"lettore": "coerenza", "cercato": "rinvio in avanti",
                                 "esito": f"udienza {f.get('valore')} anteriore alla PEC"})
        if f.get("tipo") == "sentenza" and f.get("data") and ricez and f["data"] > str(ricez):
            f["certezza"] = "da_verificare"
            d["rifiuti"].append({"lettore": "coerenza", "cercato": "data di pubblicazione",
                                 "esito": f"{f['data']} successiva alla PEC"})
    return []

def _l_atto_resto(d):
    """Quello che non e' ancora stato estratto in un lettore proprio: rinvio,
    CTU, termini, appello. Si smonta un pezzo per volta, ogni volta misurando
    sul banco di prova che i fatti letti non diminuiscano."""
    fatti = leggi_atto_monolite(d["testi"], d["quando"], d["trib"], d["xmlatto"],
                                d["rif"], d["documenti"])
    if d.get("atto") in ATTI_UO:
        via = [f for f in fatti if f.get("tipo") in FATTI_DI_SENTENZA]
        if via:
            d["rifiuti"].append({"lettore": "atto-resto", "cercato": "sentenza",
                                 "esito": "PEC di istruttoria della UO: sentenze citate ignorate"})
        fatti = [f for f in fatti if f.get("tipo") not in FATTI_DI_SENTENZA]
    return fatti
_l_atto_resto.nome, _l_atto_resto.fase, _l_atto_resto.targa = "atto-resto", 3, "2026-09-27"
LETTORI.append(_l_atto_resto)

def leggi_atto(testi, quando=None, trib=None, xmlatto=None, rif=None, documenti=None,
               oggetto=None):
    """IL COMANDO. Tiene il dossier, interroga i lettori in ordine di fase e
    restituisce i fatti timbrati. Il dossier completo - con i rifiuti e le
    targhe - resta in ULTIMO_DOSSIER e finisce in letture.json."""
    global ULTIMO_DOSSIER
    tutto = "\n".join(testo(t) for t in (testi or []) if testo(t))
    d = {"testi": testi, "quando": quando, "trib": trib or {}, "xmlatto": xmlatto or {},
         "rif": rif or {}, "documenti": list(documenti or []), "tutto": tutto,
         "oggetto": oggetto or (trib or {}).get("oggetto") or "",
         "ricez": (quando.date() if hasattr(quando, "date") else quando),
         "atto": None, "fatti": [], "rifiuti": [], "targhe": {}}
    if not tutto.strip() and not testo(d["oggetto"]).strip():
        ULTIMO_DOSSIER = d
        return []
    visti = set()
    for fn in sorted(LETTORI, key=lambda x: x.fase):
        try:
            for f in _timbra(fn(d), fn):
                k = (f.get("tipo"), testo(f.get("valore")), testo(f.get("campo")))
                if k in visti:
                    continue          # l'ha gia' detto un lettore che viene prima
                visti.add(k)
                d["fatti"].append(f)
        except Exception as e:
            d["rifiuti"].append({"lettore": fn.nome, "cercato": "lettura",
                                 "esito": f"errore: {e}"})
            err(f"lettore {fn.nome}: {e}")
        d["targhe"][fn.nome] = fn.targa
    for f in d["fatti"]:
        f["documenti"] = d["documenti"]
        n = testo(f.get("lettore")) or "?"
        CONTA_LETTORI[n] = CONTA_LETTORI.get(n, 0) + 1
    for x in d["rifiuti"]:
        k = f"{testo(x.get('lettore'))}: {testo(x.get('cercato'))} \u2014 {testo(x.get('esito'))}"
        CONTA_RIFIUTI[k] = CONTA_RIFIUTI.get(k, 0) + 1
    a = testo(d.get("atto")) or "ignoto"
    CONTA_ATTI[a] = CONTA_ATTI.get(a, 0) + 1
    ULTIMO_DOSSIER = d
    return d["fatti"]

def dossier_in_chiaro(d=None):
    """Il dossier come va su disco: solo quello che serve al registro."""
    d = d or ULTIMO_DOSSIER
    if not d:
        return None
    return {"atto": d.get("atto"), "pratica": d.get("pratica"), "ruolo": d.get("ruolo"),
            "giudice": d.get("giudice"), "fatti": d.get("fatti") or [],
            "rifiuti": d.get("rifiuti") or [], "targhe": d.get("targhe") or {},
            "letto_il": datetime.now().isoformat(timespec="seconds"),
            "intake": VERSIONE}

def leggi_atto_monolite(testi, quando=None, trib=None, xmlatto=None, rif=None,
                        documenti=None, solo_resto=False):
    """Tutti i fatti che gli atti allegati portano, in un elenco.

    `testi` sono le fonti in ordine di attendibilita': il modulo di
    cancelleria, l'XML dell'atto, il testo dei PDF, il corpo della PEC.
    `quando` e' la data di ricezione: serve per due cose che senza di essa non
    si possono dire - un rinvio va sempre avanti, e una sentenza notificata
    oggi non e' stata pubblicata domani."""
    trib = trib or {}
    xmlatto = xmlatto or {}
    rif = rif or {}
    tutto = "\n".join(testo(t) for t in (testi or []) if testo(t))
    if not tutto.strip():
        return []
    ricez = quando.date() if hasattr(quando, "date") else quando
    fatti = []

    # 0. QUELLO CHE LA DESCRIZIONE DICE IN CHIARO. Vocabolario chiuso della
    #    cancelleria: ci arriva per primo perche' e' la fonte piu' attendibile
    #    che abbiamo, e i doppioni li toglie l'unione piu' sotto.
    for x in fatti_dalla_descrizione(trib.get("descrizione"), ricez):
        fatti.append(x)

    # 1. IL RINVIO. Il fatto piu' frequente e il piu' urgente: e' la data che
    #    regge tutte le scadenze del fascicolo.
    ud, frase_ud = udienza_annunciata(
        [trib.get("descrizione"), trib.get("oggetto"), tutto], ricez)
    if ud:
        fatti.append(_fatto("udienza", ud.isoformat(),
                            f"udienza {ud.strftime('%d/%m/%Y')}", frase_ud,
                            "alta" if trib.get("descrizione") else "da_verificare"))

    # 2. LA SENTENZA. Il numero puo' stare nell'oggetto (lo trova riferimenti())
    #    o soltanto dentro il PDF: qui si cerca anche li'.
    num = rif.get("sentenza")
    frase_s = None
    if not num:
        m = RE_SENT_TESTO.search(tutto)
        if m:
            num = f"{int(m.group(1))}/{m.group(2)}"
            frase_s = _frase(tutto, m.start(), m.end())
    if num:
        # la data di pubblicazione sta scritta: e' meglio della data della PEC
        pub, frase_p = _data_dopo(tutto, r"(?:pubblicat\w+|depositat\w+)\s+(?:il|in\s+data)")
        fatti.append(_fatto("sentenza", num, f"sentenza {num}", frase_s or frase_p,
                            "alta" if frase_s or rif.get("sentenza") else "da_verificare",
                            data=(pub or ricez).isoformat() if (pub or ricez) else None,
                            data_certa=bool(pub)))

    # 3. L'ESITO. Si legge nel dispositivo, che sta in fondo all'atto. Se il
    #    numero di sentenza non c'e', un esito da solo vale comunque: dice che
    #    il fascicolo e' definito.
    for chiave, etichetta, rx in ESITI:
        m, frase_e = _cerca(tutto, rx)
        if m:
            fatti.append(_fatto("esito", chiave, etichetta, frase_e,
                                "da_verificare" if not num else "alta"))
            break

    # 4. LE SPESE. Tre forme, e sono alternative fra loro.
    m, frase_sp = _cerca(tutto, r"spes\w+[^.\n]{0,40}?(?:integralment\w+\s+)?compensat|"
                                r"compens\w+\s+(?:integralmente\s+)?le\s+spes")
    if m:
        fatti.append(_fatto("spese", "compensate", "spese compensate", frase_sp))
    else:
        m, frase_sp = _cerca(tutto, r"spes\w+[^.\n]{0,30}?irripetibil")
        if m:
            fatti.append(_fatto("spese", "irripetibili", "spese irripetibili", frase_sp))
        else:
            m = RE_SPESE_EURO.search(tutto)
            if m:
                imp = m.group(1) or m.group(2)
                fatti.append(_fatto("spese", imp, f"spese liquidate \u20ac {imp}",
                                    _frase(tutto, m.start(), m.end()), "da_verificare"))

    # 5. LA CTU. Nomina, chi e', quando giura e quando deve depositare: sono
    #    quattro date diverse e nessuna delle quattro e' l'udienza.
    m, frase_c = _cerca(tutto, r"nomin\w+\s+(?:il\s+|un\s+)?(?:C\.?T\.?U\.?|"
                               r"consulent\w+\s+tecnic\w+)|disponend\w+\s+C\.?T\.?U|"
                               r"si\s+dispone\s+(?:la\s+)?C\.?T\.?U")
    if m:
        nome = RE_CTU_NOME.search(tutto)
        fatti.append(_fatto("ctu", (nome.group(1).strip() if nome else True),
                            "CTU" + (f": {nome.group(1).strip()}" if nome else " disposta"),
                            frase_c))
    giura, frase_g = _data_dopo(
        tutto, r"(?:giuramento|conferiment\w+\s+dell['\u2019]\s*incaric\w+|"
               r"inizio\s+delle\s+operazioni)[^\n0-9]{0,40}?(?:al|il|per\s+il|in\s+data)", ricez)
    if giura:
        fatti.append(_fatto("ctu_giuramento", giura.isoformat(),
                            f"giuramento CTU {giura.strftime('%d/%m/%Y')}", frase_g))
    dep, frase_d = _data_dopo(
        tutto, r"deposit\w+\s+(?:della\s+|dell['\u2019]\s*)?(?:relazion\w+|elaborat\w+|perizi\w+)"
               r"[^\n0-9]{0,40}?(?:entro\s+(?:il\s+)?|il|in\s+data)", ricez)
    if dep:
        fatti.append(_fatto("ctu_deposito", dep.isoformat(),
                            f"relazione CTU entro il {dep.strftime('%d/%m/%Y')}", frase_d))

    # 6. I TERMINI concessi dal giudice: note, memorie, documenti. Sono
    #    adempimenti nostri, e scadono prima dell'udienza.
    term, frase_t = _data_dopo(
        tutto, r"(?:note|memori\w+|osservazion\w+|document\w+)\s+(?:scritt\w+\s+)?"
               r"(?:autorizzat\w+\s+)?(?:da\s+deposit\w+\s+)?entro\s+(?:il\s+)?", ricez)
    if term:
        fatti.append(_fatto("termine", term.isoformat(),
                            f"termine per note {term.strftime('%d/%m/%Y')}", frase_t,
                            "da_verificare"))

    # 7. IL PASSAGGIO IN APPELLO. Non aggiorna questo fascicolo: ne apre un
    #    altro. Va detto, non applicato da solo.
    m = RE_APPELLO.search(tutto)
    if m and (rif.get("grado") != "appello"):
        fatti.append(_fatto("appello", True, "impugnazione: apre il secondo grado",
                            _frase(tutto, m.start(), m.end()), "da_verificare"))

    # 8. IL GIUDICE, quando l'atto lo nomina e il modulo no: serve a capire di
    #    chi e' il fascicolo.
    gd = trib.get("giudice") or xmlatto.get("giudice")
    if not gd:
        m = RE_GIUDICE_TESTO.search(tutto)
        if m:
            gd = m.group(1).strip()
            fatti.append(_fatto("giudice", gd, f"giudice: {gd}",
                                _frase(tutto, m.start(), m.end()), "da_verificare"))

    # 9. I RIFERIMENTI che l'atto porta e il fascicolo forse non ha: un ruolo
    #    trovato nel PDF vale quanto uno trovato nell'oggetto.
    r2 = riferimenti(tutto[:30000]) or {}
    for chiave, eti in (("rg", "numero di ruolo"), ("sisco", "codice SISCO"),
                        ("spese_legali", "pratica spese legali")):
        if r2.get(chiave) and not rif.get(chiave):
            fatti.append(_fatto("riferimento", r2[chiave],
                                f"{eti} {r2[chiave]}", None, "da_verificare",
                                campo=chiave))

    # LA DIAGNOSTICA CHE MANCAVA. Una descrizione che non produce nessun fatto
    # e' una frase della cancelleria che non sappiamo leggere, e finora non lo
    # diceva nessuno: le due sentenze e la CTU non lette stavano in archivio da
    # mesi senza che nulla lo segnalasse. Qui si contano per FORMA - i numeri
    # diventano N - cosi' nel referto si leggono venti righe invece di mille, e
    # il vocabolario cresce su quello che arriva davvero.
    d = testo(trib.get("descrizione"))
    if d and not fatti:
        forma = re.sub(r"\d+", "N", d)
        forma = re.sub(r"\s+", " ", forma).strip()[:90]
        DESCRIZIONI_MUTE[forma] = DESCRIZIONI_MUTE.get(forma, 0) + 1

    for f in fatti:
        f["documenti"] = list(documenti or [])
    return fatti

def fatti_in_chiaro(novita):
    """Una riga di log: cosa ha imparato il fascicolo da questa PEC."""
    if not novita:
        return "nessun fatto negli allegati"
    return ", ".join(testo(f.get("etichetta")) +
                     ("?" if f.get("certezza") != "alta" else "")
                     for f in novita[:6])

def chiave_vera(k):
    return bool(re.match(r"^(mid|hash|vecchia|prova):", testo(k)))

def chiave_comunicazione(v, cartella=""):
    """Chi e' questa comunicazione: la chiave del messaggio se c'e', altrimenti
    una fatta dalla voce stessa. Le voci vecchie portano una chiave finta
    ("parte" su 288 fascicoli): trattarla come vera faceva cancellare una
    comunicazione quando ne arrivava un'altra con la stessa parola, e nel
    registro faceva leggere a tutte i fatti della prima. Stessa regola del
    registro (chiaveComunicazione)."""
    if chiave_vera(v.get("chiave")):
        return v["chiave"]
    base = "|".join([testo(cartella), testo(v.get("ricevuta")),
                     testo(v.get("oggetto_reale") or v.get("oggetto"))])
    return "vecchia:" + hashlib.sha256(base.encode("utf-8", "ignore")).hexdigest()[:24]

def aggiungi_comunicazione(fasc, voce):
    """Registra una comunicazione nel fascicolo senza duplicarla: la chiave del
    messaggio identifica il fatto. Ripulisce anche i doppioni gia' presenti,
    cosi' una rilettura non lascia tracce."""
    p = fasc / "comunicazioni.json"
    reg = leggi_json(p, {"comunicazioni": []})
    elenco = reg.get("comunicazioni") or []
    voce["chiave"] = chiave_comunicazione(voce, fasc.name)
    chiave = voce["chiave"]

    visti, puliti, tolti = set(), [], 0
    for v in elenco:
        if not chiave_vera(v.get("chiave")):
            v["chiave_scritta"] = v.get("chiave")
            v["chiave"] = chiave_comunicazione(v, fasc.name)
        k = v["chiave"]
        if k in visti:
            tolti += 1
            continue
        visti.add(k)
        puliti.append(v)

    puliti = [v for v in puliti if (v.get("chiave") or "") != chiave] if chiave else puliti
    puliti.append(voce)
    reg["comunicazioni"] = puliti
    reg["aggiornato"] = datetime.now().isoformat(timespec="seconds")
    scrivi_json(p, reg)
    scrivi_lettura(fasc, chiave, voce)
    return tolti

def scrivi_lettura(fasc, chiave, voce=None):
    """letture.json: il dossier di lettura, una voce per PEC. E' la parola
    definitiva di chi legge, e l'unico punto di passaggio verso il registro -
    che legge questo file e decide cosa applicare. Qui non si scrive niente
    nei fascicoli: le due piramidi non si toccano i campi a vicenda."""
    d = dossier_in_chiaro()
    if not d or not chiave:
        return
    p = Path(fasc) / "letture.json"
    reg = leggi_json(p, {"letture": {}})
    reg.setdefault("letture", {})
    d["oggetto"] = testo((voce or {}).get("oggetto_reale") or (voce or {}).get("oggetto"))[:200]
    d["ricevuta"] = testo((voce or {}).get("ricevuta"))[:19]
    reg["letture"][chiave] = d
    reg["aggiornato"] = datetime.now().isoformat(timespec="seconds")
    scrivi_json(p, reg)
    LETTURE["scritte"] += 1

LANCI = BASE / "lanci.json"
DIARIO = BASE / "diario_lanci.txt"

def diario(esito, dettaglio=""):
    """Una riga per lancio, in fondo, e non si cancella mai.

    Il referto viene riscritto a ogni giro: con l'esecuzione automatica vedi
    solo l'ultimo, e un giro fallito alle tre di notte non lo sai mai. Qui
    invece resta traccia di TUTTI, e soprattutto di quelli andati male: e'
    l'unico modo per accorgersi che l'automatismo si e' fermato.

    Una riga sola, leggibile senza aprire niente:
      2026-09-14 22:00  OK        31157 letti, 8520 lavorati, 67 cartelle
      2026-09-15 22:00  FALLITO   collegamento a Outlook: profilo protetto
    """
    try:
        riga = (datetime.now().strftime("%Y-%m-%d %H:%M") + "  " +
                esito.ljust(11) + " " + testo(dettaglio)[:150] + "\n")
        with open(DIARIO, "a", encoding="utf-8") as f:
            f.write(riga)
    except Exception:
        pass          # il diario non deve mai far cadere il programma
LANCIO = {}

def apri_lancio(cfg, args):
    """Ogni corsa lascia traccia di cosa ha creato: senza questo non si puo'
    tornare indietro, e una prova che non si puo' annullare non e' una prova."""
    global LANCIO
    LANCIO = {
        "id": datetime.now().strftime("%Y%m%d_%H%M%S"),
        "iniziato": datetime.now().isoformat(timespec="seconds"),
        "finestra": (f"{cfg.get('dal')} - {cfg.get('al')}" if cfg.get("dal")
                     else f"ultimi {cfg.get('giorni_indietro')} giorni"),
        "cartelle_transito": [],     # create in IN ARRIVO
        "fascicoli_creati": [],      # cartelle nuove dentro l'archivio dei fascicoli
        "code": [],                  # sottocartelle COMUNICAZIONI e righe di registro
        "unioni": [],                # file aggiunti a cartelle gia' esistenti
        "chiavi": [],                # messaggi elaborati
    }
    return LANCIO

def chiudi_lancio(esito):
    if not LANCIO:
        return
    LANCIO["concluso"] = datetime.now().isoformat(timespec="seconds")
    LANCIO.update(esito or {})
    storico = leggi_json(LANCI, {"lanci": []})
    storico.setdefault("lanci", []).append(LANCIO)
    storico["lanci"] = storico["lanci"][-20:]     # gli ultimi venti bastano
    scrivi_json(LANCI, storico)
    log(f"lancio {LANCIO['id']} registrato: "
        f"{len(LANCIO['cartelle_transito'])} cartelle di transito, "
        f"{len(LANCIO['fascicoli_creati'])} fascicoli creati, "
        f"{len(LANCIO['code'])} comunicazioni accodate")
    log(f"per annullarlo:  python {Path(__file__).name} --annulla {LANCIO['id']}")

def annulla_lancio(quale):
    """Rimette le cose come stavano: toglie le cartelle create, le comunicazioni
    accodate e i file aggiunti. Non tocca nulla che non abbia creato lui."""
    storico = leggi_json(LANCI, {"lanci": []})
    lanci = storico.get("lanci") or []
    if not lanci:
        err("nessun lancio registrato da annullare")
        return 1
    if quale in (None, "", "ultimo"):
        l = lanci[-1]
    else:
        l = next((x for x in lanci if x.get("id") == quale), None)
        if not l:
            err(f"lancio '{quale}' non trovato. Disponibili:")
            for x in lanci[-8:]:
                err(f"   {x['id']}  {x.get('finestra','')}  "
                    f"{len(x.get('cartelle_transito',[]))} cartelle")
            return 1

    log(f"annullo il lancio {l['id']} del {l.get('iniziato','')[:16]} ({l.get('finestra','')})")
    tolti = {"file": 0, "code": 0, "cartelle": 0, "fascicoli": 0, "righe": 0}

    # 1. i file aggiunti a cartelle preesistenti: solo quelli, non gli altri
    for u in l.get("unioni", []):
        for nome in u.get("file", []):
            p = Path(u["cartella"]) / nome
            try:
                if p.exists():
                    p.unlink()
                    tolti["file"] += 1
            except Exception as e:
                err(f"non rimosso {p.name}: {e}")

    # 2. le sottocartelle COMUNICAZIONI e le righe nei registri dei fascicoli
    per_fascicolo = {}
    for c in l.get("code", []):
        per_fascicolo.setdefault(c["fascicolo"], []).append(c)
    for fascicolo, voci in per_fascicolo.items():
        chiavi = {v.get("chiave") for v in voci if v.get("chiave")}
        for v in voci:
            if v.get("box") and rimuovi(v["box"]):
                tolti["code"] += 1
        p = Path(fascicolo) / "comunicazioni.json"
        if p.exists():
            reg = leggi_json(p, {})
            prima = len(reg.get("comunicazioni") or [])
            reg["comunicazioni"] = [x for x in (reg.get("comunicazioni") or [])
                                    if x.get("chiave") not in chiavi]
            tolti["righe"] += prima - len(reg["comunicazioni"])
            if reg["comunicazioni"]:
                scrivi_json(p, reg)
            else:
                p.unlink(missing_ok=True)
        try:
            coda = Path(fascicolo) / "COMUNICAZIONI"
            if coda.exists() and not any(coda.iterdir()):
                coda.rmdir()
        except Exception:
            pass

    # 3. le cartelle di transito create dal lancio
    for d in l.get("cartelle_transito", []):
        if Path(d).exists() and rimuovi(d):
            tolti["cartelle"] += 1

    # 4. le cartelle di fascicolo create, solo se vuote: se qualcuno ci ha
    #    messo qualcosa dentro non si tocca
    for d in l.get("fascicoli_creati", []):
        p = Path(d)
        try:
            if not p.exists():
                continue
            resto = [x for x in p.rglob("*") if x.name != ".fascicolo.json"]
            if resto:
                err(f"cartella {p.name} NON rimossa: contiene {len(resto)} elementi")
                continue
            if rimuovi(p):
                tolti["fascicoli"] += 1
            padre = p.parent
            if padre.exists() and not any(padre.iterdir()):
                padre.rmdir()
        except Exception as e:
            err(f"non rimossa {p.name}: {e}")

    # 5. i messaggi tornano da elaborare
    arch = leggi_json(ARCHIVIO, {"visti": {}})
    visti = arch.get("visti") or {}
    for k in l.get("chiavi", []):
        visti.pop(k, None)
    arch["visti"] = visti
    arch["copie"] = {}
    scrivi_json(ARCHIVIO, arch)

    storico["lanci"] = [x for x in lanci if x.get("id") != l["id"]]
    scrivi_json(LANCI, storico)

    log("")
    log(f"ANNULLATO il lancio {l['id']}:")
    log(f"  {tolti['cartelle']} cartelle di transito rimosse")
    log(f"  {tolti['fascicoli']} cartelle di fascicolo rimosse (solo quelle rimaste vuote)")
    log(f"  {tolti['code']} sottocartelle COMUNICAZIONI rimosse")
    log(f"  {tolti['righe']} righe tolte dai registri dei fascicoli")
    log(f"  {tolti['file']} file tolti da cartelle preesistenti")
    log(f"  {len(l.get('chiavi', []))} messaggi tornati da elaborare")
    return 0

RE_IDENTIFICATIVO = re.compile(
    r"identificativo\s+messaggio\s*:?\s*([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+)", re.I)

def identificativo_originale(esito, corpo):
    """Ogni PEC ha un identificativo che il gestore assegna una volta sola.
    L'inoltro lo porta dentro, nel corpo e nel daticert: e' il modo sicuro
    per capire che quella comunicazione l'abbiamo gia' vista, anche se
    l'inoltro arriva mesi dopo con un altro oggetto."""
    ids = []
    for d in (esito.get("daticert_tutti") or []):
        for k in ("identificativo", "msgid", "message_id"):
            if d.get(k):
                ids.append(testo(d[k]).strip().strip("<>"))
    for m in RE_IDENTIFICATIVO.finditer(testo(corpo)[:20000]):
        ids.append(m.group(1).strip().strip("<>"))
    # il piu' antico e' l'originale: gli inoltri aggiungono i propri in coda
    visti, fuori = set(), []
    for x in ids:
        if x and x.lower() not in visti:
            visti.add(x.lower())
            fuori.append(x)
    return fuori

RE_INOLTRO = re.compile(r"^\s*(?:i|r|re|rif|fw|fwd|inoltro|inoltra)\s*:", re.I)

def e_inoltro_interno(mittente, destinatari, oggetto, esito):
    """Un collega che gira una PEC a un altro collega: mittente e destinatari
    sono tutti dell'Istituto, e l'oggetto porta il prefisso di inoltro.
    Non e' una comunicazione nuova, e' la stessa che gira."""
    def interno(x):
        return bool(re.search(r"@inps\.it|@postacert\.inps\.gov\.it", testo(x), re.I))
    mitt_int = interno(mittente)
    dest = testo(destinatari)
    dest_int = interno(dest) and not re.search(
        r"@(?!inps\.it|postacert\.inps\.gov\.it)[a-z0-9.-]+\.[a-z]{2,}", dest, re.I)
    prefisso = bool(RE_INOLTRO.match(testo(oggetto))) or \
        len(re.findall(r"\b(?:fwd?|i|r|re|rif)\s*:", testo(oggetto), re.I)) >= 1
    return mitt_int and dest_int and prefisso

def e_pec_da_fascicolo(esito, dati, allegati, mittente, oggetto, corpo):
    """LA REGOLA. Restituisce (True, None) se la PEC merita un fascicolo,
    oppure (False, motivo) se e' una comunicazione che non ne merita uno.

    Una PEC fa nascere un fascicolo quando porta ALMENO UNO di questi segni,
    che solo un atto vero puo' avere:

      1. l'attestazione del gestore con un mittente ESTERNO all'Istituto
         (daticert.xml: e' la prova che qualcuno ci ha notificato qualcosa)
      2. la busta del processo telematico (XML dell'atto)
      3. il modulo di cancelleria del tribunale
      4. un allegato che e' un atto: ricorso, relata, notifica ex legge 53,
         decreto di fissazione, DFU, procura, citazione
      5. un riferimento processuale COMPLETO: numero di ruolo o SISCO

    E non ne merita uno quando:

      - mittente e destinatari sono tutti dell'Istituto (posta fra colleghi)
      - non ha nessuno dei cinque segni: e' una nota di servizio
      - ha solo un nominativo, senza atti e senza riferimenti

    Il senso: un ufficio legale riceve molta posta, ma un ricorso lascia
    tracce riconoscibili. Dove non ce ne sono, aprire un fascicolo significa
    sommergere quelli veri.
    """
    testo_tot = (esito.get("testo") or "") + " " + testo(corpo)

    # --- 1. attestazione con mittente esterno ---
    mitt = testo(dati.get("mittente_pec") or mittente)
    da_dentro = bool(re.search(r"@inps\.it|@postacert\.inps\.gov\.it", mitt, re.I))
    if (esito.get("daticert_tutti") or dati.get("mittente_pec")) and not da_dentro:
        return True, None

    # --- 2. busta telematica ---
    if esito.get("xml_atto"):
        return True, None

    # --- 3. modulo di cancelleria ---
    if dati.get("da_modulo") or dati.get("ufficio_giudiziario"):
        return True, None

    # --- 4. un allegato che e' un atto ---
    RE_ATTO = re.compile(
        r"ricors|relata|relazione\s+di\s+notific|notific.*(?:53|art\.?\s*3\s*bis)|"
        r"decret.*fissaz|fissazione.*udienz|\bdfu\b|procura|citazion|"
        r"comparsa|memoria|verbale.*udienz|sentenz|ordinanz|precett", re.I)
    for a in allegati:
        # gli allegati arrivano come stringhe (il nome del file), non come
        # oggetti: trattarli come dizionari significava non vederli mai
        nome_a = a if isinstance(a, str) else (a.get("nome") if hasattr(a, "get") else "")
        if RE_ATTO.search(testo(nome_a)):
            return True, None
    if RE_ATTO.search(testo(oggetto)):
        return True, None

    # --- 5. un riferimento processuale, anche solo nell'oggetto ---
    if dati.get("rg") or dati.get("sisco"):
        return True, None
    # "COMUNICAZIONE 2807/2026/LAV" porta il ruolo nel titolo: e' un
    # riferimento a tutti gli effetti, anche se non e' stato estratto
    if re.search(r"\b\d{1,6}\s*/\s*(?:19|20)\d{2}\s*/\s*(?:LAV|CIV|PREV)\b|"
                 r"\bR\.?\s*G\.?\s*n?\.?\s*\d{1,6}\s*/\s*(?:19|20)?\d{2}\b|"
                 r"\b\d{3,4}\s*/\s*\d{2}\s*/\s*(?:CO|AO|CV)\b", testo(oggetto), re.I):
        return True, None

    # --- nessun segno: si guarda chi scrive ---
    if da_dentro:
        return False, "posta interna: mittente dell'Istituto e nessun atto allegato"
    if not allegati:
        return False, "nessun allegato e nessun riferimento processuale"
    return False, ("nessun atto riconoscibile fra gli allegati e nessun "
                   "numero di ruolo")

# Oggetti che non identificano nulla: "Notificazione ai sensi della legge 53"
# ne arrivano decine al giorno, da avvocati diversi e per cause diverse.
RE_OGGETTO_GENERICO = re.compile(
    r"^(?:notificazione|notifica|notificazioni)\b[^0-9]{0,20}"
    r"(?:ai\s+sensi|ex\b|a\s+mezzo|in\s+proprio)|"
    r"^(?:posta\s+certificata|comunicazione|trasmissione|invio|notifica)\s*:?\s*$|"
    r"^atti\s+(?:allegati|in\s+notifica)|^deposito\s+telematico|"
    r"^relata\b|^relazione\s+di\s+notific", re.I)

def segna_copia(firma, dove):
    """Ogni volta che una comunicazione prende posto, la sua firma resta qui:
    la prossima PEC con la stessa firma e' una copia. Il registro delle firme
    del lancio serve solo al referto, per distinguere le copie di oggi da
    quelle di comunicazioni gia' archiviate."""
    COPIE[firma] = dove
    FIRME_LANCIO.add(firma)

def firma_comunicazione(oggetto_reale, quando, rif, extra=None):
    """La stessa PEC arriva alla casella dedicata E a quella dell'avvocato
    istruttore: sono messaggi distinti (Message-ID diverso) ma una sola
    comunicazione. La riconosco da oggetto + giorno + riferimenti.

    ATTENZIONE: se l'oggetto e' una formula di rito e non ci sono riferimenti,
    due notifiche diverse dello stesso giorno avrebbero la stessa firma e la
    seconda sparirebbe. In quel caso serve qualcosa che le distingua davvero:
    l'ora esatta e il mittente."""
    ogg = re.sub(r"\s+", " ", unicodedata.normalize("NFKC", testo(oggetto_reale)).upper()).strip()
    chiavi = "|".join(f"{k}={v}" for k, v in sorted((rif or {}).items()))
    generico = not chiavi and (RE_OGGETTO_GENERICO.search(ogg) or len(ogg) < 24)
    if generico:
        # ora al minuto e mittente: due PEC distinte non coincidono mai.
        #
        # Il mittente va normalizzato: Outlook lo espone a volte come indirizzo
        # SMTP, a volte nella forma interna di Exchange
        # (/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP/...). Sono la stessa
        # casella, ma come stringhe sono diverse, e la stessa PEC finiva con due
        # firme diverse: 48 cartelle doppie che nessun controllo intercettava.
        # Della forma interna non si ricava l'indirizzo, quindi si scarta: a
        # distinguere restano oggetto e ora al minuto, che con la stessa
        # comunicazione coincidono e con due diverse no.
        # Il mittente NON entra nella firma. Outlook lo espone a volte come
        # indirizzo SMTP e a volte nella forma interna di Exchange
        # (/O=EXCHANGELABS/OU=...): la stessa PEC prendeva due firme diverse e
        # nascevano due cartelle. Dalla forma interna l'indirizzo non si
        # ricava, quindi non c'e' modo di riconciliarle.
        #
        # Verificato prima di toglierlo: sui 279 gruppi con lo stesso oggetto
        # e lo stesso minuto, riferimenti e nominativi coincidono sempre.
        # Nessuno di quei gruppi e' fatto di comunicazioni diverse, quindi
        # oggetto piu' minuto basta a distinguerle.
        base = f"{ogg}|{str(quando)[:16]}"
    else:
        base = f"{ogg}|{str(quando)[:10]}|{chiavi}"
    # Quale delle due regole ha deciso: il referto lo dice, perche' le due non
    # sono affidabili allo stesso modo. Oggetto + giorno + riferimenti
    # identifica una comunicazione; oggetto + minuto e' un ripiego per gli
    # oggetti di rito senza riferimenti, ed e' la regola che puo' confondere
    # due PEC diverse arrivate nello stesso minuto.
    firma_comunicazione.regola = ("oggetto e minuto, nessun riferimento"
                                  if generico else
                                  "oggetto, giorno e riferimenti")
    return hashlib.sha256(base.encode()).hexdigest()[:20]


    """La stessa PEC arriva alla casella dedicata E a quella dell'avvocato
    istruttore: sono messaggi distinti (Message-ID diverso) ma una sola
    comunicazione. La riconosco da oggetto + giorno + riferimenti."""
    ogg = re.sub(r"\s+", " ", unicodedata.normalize("NFKC", testo(oggetto_reale)).upper()).strip()
    giorno = str(quando)[:10]
    chiavi = "|".join(f"{k}={v}" for k, v in sorted((rif or {}).items()))
    return hashlib.sha256(f"{ogg}|{giorno}|{chiavi}".encode()).hexdigest()[:20]

def _stesso_messaggio(cartella, chiave):
    """La cartella e' gia' di questo messaggio? Si legge la chiave dalla sua
    scheda: e' l'unico dato che non cambia fra un giro e l'altro."""
    if not chiave:
        return False
    try:
        d = json.loads((cartella / "scheda_pec.json").read_text(encoding="utf-8"))
    except Exception:
        return False
    return testo(d.get("chiave")) == testo(chiave)

def prossima_cartella(radice, quando, tipo=None, rif=None, oggetto=""):
    """Il nome deve dire cosa c'e' dentro senza doverla aprire:
    2026-08-18 pagamento RG 578-2021 - Campisi Corradina spese legali"""
    try:
        g = f"{quando.year:04d}-{quando.month:02d}-{quando.day:02d}"
    except Exception:
        g = datetime.now().strftime("%Y-%m-%d")

    pezzi = [g]
    if tipo and tipo != "da_classificare":
        pezzi.append(tipo.replace("_", " "))
    r = rif or {}
    for etichetta, chiave in (("SISCO", "sisco"), ("RG", "rg"),
                              ("sent", "sentenza"), ("com", "comunicazione"),
                              ("spese", "spese_legali")):
        if r.get(chiave):
            pezzi.append(f"{etichetta} {str(r[chiave]).replace('/', '-')}")
            break
    coda = re.sub(r"\s+", " ", testo(oggetto)).strip()
    if len(coda) > 46:
        taglio = coda[:46].rsplit(" ", 1)[0]
        coda = taglio if len(taglio) > 20 else coda[:46]
    if coda:
        pezzi.append(coda)
    nome = pulisci_nome(" - ".join(pezzi), 120)

    # Prima di creare: la cartella di questo messaggio puo' essere gia' stata
    # spostata in una delle sottocartelle di accantonamento. Cercando solo in
    # radice la si ricreava, e la PEC compariva due volte - una in
    # DA ESAMINARE e una fuori.
    for sotto in ("ALTRI GIUDICI", "DA ESAMINARE", "ESECUZIONI"):
        q = radice / sotto / nome
        if q.exists() and _stesso_messaggio(q, prossima_cartella.chiave_corrente):
            return q

    p = radice / nome
    n = 2
    while p.exists():
        # Se la cartella esiste ed e' dello STESSO messaggio, si riusa: con
        # --rileggi ogni messaggio viene rilavorato, e creando "(2)" accanto
        # alla vecchia si duplicava l'intero transito. La cartella si
        # riconosce dalla chiave del messaggio scritta nella sua scheda, non
        # dal nome.
        if _stesso_messaggio(p, prossima_cartella.chiave_corrente):
            return p
        p = radice / f"{nome} ({n})"
        n += 1
    return p

prossima_cartella.chiave_corrente = None

# --------------------------------------------------------------------------
def elabora_messaggio(m, cart, k, motivo, args, dest_radice, visti_mail, referto):
    """Lavorazione di una singola PEC. Se qui va storto qualcosa il chiamante
    conta l'errore e prosegue: una PEC malata non ferma le altre."""
    ogg = testo(prop(m, "Subject"))
    mit = mittente_vero(m)
    quando = prop(m, "ReceivedTime", None)

    allegati = []
    try:
        for j in range(1, m.Attachments.Count + 1):
            allegati.append(testo(prop(m.Attachments.Item(j), "FileName")))
    except Exception:
        pass

    reale = oggetto_reale(ogg)
    corpo_avviso = testo(prop(m, "Body"))
    # Il modulo delle cancellerie sta nel corpo della PEC: si legge subito,
    # perche' serve anche agli avvisi che non hanno allegati.
    trib = leggi_comunicazione_tribunale(corpo_avviso)
    scheda_trib = {k: v for k, v in trib.items() if v not in (None, "", False)} or None
    portale = leggi_avviso_portale(corpo_avviso)

    tipo = tipo_di_pec(reale, allegati, TIPI)
    rif = riferimenti(reale)
    if trib.get("rg"):
        rif["rg"] = trib["rg"]
    fase = fase_ctu(reale, allegati, FASI) \
        if (tipo in ("comunicazione_cancelleria", "ctu") or rif.get("comunicazione")) else None
    # Il modulo "(Sede: X - Pratica: N - NOMINATIVO)" e' riconoscibile quanto
    # il numero di comunicazione: e' la cancelleria che scrive, e finiva fra
    # le "da classificare" solo perche' il numero non c'era.
    if (rif.get("comunicazione") or RE_MODULO.search(testo(reale))) \
       and tipo in ("da_classificare", "da_leggere"):
        tipo = "comunicazione_cancelleria"
    nota_dest = casella_di_notifica(m, "", CASELLE) if args.ricognizione else None
    voce_referto = {"data": str(quando)[:16], "mittente": mit,
                    "cartella": testo(prop(cart, "Name")),
                    "oggetto": reale[:90], "motivo": motivo,
                    "tipo": tipo, "fase": fase, "rif": rif, "allegati": allegati,
                    "notifica": (nota_dest or {}).get("notifica_valida"),
                    "fonte_notifica": None, "daticert": False}
    referto.append(voce_referto)

    if args.ricognizione:
        # In ricognizione gli allegati non si aprono: qui i lettori ricevono
        # soltanto oggetto e corpo. Vale la pena interrogarli comunque - la
        # trasmissione della UO e l'identita' della pratica stanno tutte
        # nell'oggetto - ma i conteggi del referto NON sono confrontabili con
        # quelli di un'elaborazione, dove i PDF ci sono. Prima di questa
        # riga i lettori non venivano interrogati affatto, e il referto
        # contava un atto solo su milletrecento messaggi.
        # IL MODULO VA PASSATO. Era letto a riga 3740 e poi buttato: qui andava
        # un None al posto di trib, e il lettore della descrizione - che e' la
        # frase in cui la cancelleria dice se sta comunicando una sentenza, un
        # rinvio o una CTU - non vedeva niente. Nel referto del 01/10: 223 atti
        # su 274 "ignoti", 238 "numero di pratica non trovato", e la sezione
        # delle descrizioni non lette vuota non perche' fossero tutte lette, ma
        # perche' nessuna era stata nemmeno guardata. La ricognizione serve a
        # guardare prima di scrivere: era l'unica cosa che non guardava.
        try:
            leggi_atto([reale, testo(prop(m, "Body"))[:20000]],
                       quando, trib, None, rif, None, oggetto=reale)
        except Exception as e:
            err(f"letture in ricognizione: {e}")
        return True

    # gli avvisi del Portale non portano documenti: creare una cartella vuota
    # per ognuno significa 900 cartelle da aprire per niente. Vanno solo in elenco.
    #
    # "Non porta documenti" non vuol dire "non ha allegati": un messaggio di un
    # collega con tre image00N.png della firma in calce ha tre allegati e zero
    # documenti. Contandoli, il filtro sulla provenienza interna non scattava e
    # nasceva una cartella con dentro solo le immagini di una firma.
    # allegati_documentali esisteva gia' e scarta proprio quelle: qui la si usa.
    def _e_documento(n):
        n = testo(n).lower()
        if re.match(r"^image\d+\.", n):
            return False          # firma in calce
        if n == "daticert.xml":
            return False          # c'e' su ogni PEC: non e' un documento
        # La busta: Outlook a volte la rinomina "postacert.eml (177 KB).msg".
        # Finisce in .msg e non in .eml, e non veniva riconosciuta come
        # documento: la PEC diventava un avviso e la busta non si apriva mai.
        # e_la_busta sa gia' riconoscerla, ma piu' avanti - qui non ci si
        # arrivava.
        if e_la_busta(n) or n.endswith(".msg"):
            return True
        # l'xml dell'atto telematico conta: certe comunicazioni non hanno altro
        return n.endswith(DOCUMENTALI) or n.endswith(".xml")
    doc_veri = [n for n in allegati if _e_documento(n)]
    solo_avviso = not doc_veri and SOLO_CON_DOCUMENTI
    firma = firma_comunicazione(reale, quando, rif, mit)
    regola_firma = getattr(firma_comunicazione, "regola", "?")

    # L'inoltro interno si riconosce da mittente e destinatari, senza aprire
    # la busta: e' un controllo che si puo' fare subito.
    if SCARTA_INOLTRI and e_inoltro_interno(mit, str(prop(m, "To", "")) + " " +
                                            str(prop(m, "CC", "")), ogg, None):
        visti_mail[k] = {"cartella": None, "oggetto": ogg[:120],
                         "ricevuta": str(quando)[:19], "documenti": 0,
                         "scartata": "inoltro interno fra colleghi: la comunicazione "
                                     "originale arriva per conto suo"}
        voce_referto["inoltro_interno"] = True
        log(f"    x inoltro interno  ({reale[:46]})")
        return "inoltro"

    gemella = COPIE.get(firma)
    if gemella == "__avviso__":
        # Un avviso senza documenti non puo' rendere copia una PEC che i
        # documenti li ha: sarebbe come perdere l'atto per via della ricevuta.
        # Vale come copia solo se anche questa e' senza allegati.
        if allegati:
            gemella = None
            COPIE.pop(firma, None)
        else:
            visti_mail[k] = {"cartella": None, "oggetto": ogg[:120],
                             "ricevuta": str(quando)[:19], "documenti": 0, "copia_di": firma}
            voce_referto["copia_perche"] = regola_firma
            voce_referto["copia_regola"] = regola_firma
            voce_referto["copia_stesso_lancio"] = firma in FIRME_LANCIO
            voce_referto["copia_di"] = testo(gemella)[:70]
            # Chiave STABILE: "copia:" + firma identifica la comunicazione,
            # non il singolo messaggio. Cosi' ogni arrivo successivo della
            # stessa PEC trova la riga esistente e non ne crea una nuova.
            if DB_ATTIVO:
                try:
                    DB.scrivi_copia_archiviata({
                        "chiave": "copia:" + firma,
                        "ricevuta": str(quando)[:19],
                        "oggetto": ogg,
                        "oggetto_reale": reale,
                        "mittente": mit,
                        "mittente_pec": mit,
                        "tipo_pec": tipo,
                        "fase_ctu": fase,
                        "riferimenti": rif,
                        "novita": [],
                    })
                except Exception as e:
                    err(f"scrittura db della copia (avviso): {e}")
            salva_archivio({"visti": visti_mail, "copie": COPIE,
                            "aggiornato": datetime.now().isoformat(timespec="seconds")})
            log(f"    = copia di avviso  ({testo(prop(cart, 'Name'))})  {reale[:44]}")
            return "copia"
    if gemella:
        # stessa comunicazione gia' lavorata: niente cartella nuova, solo la traccia
        p = Path(gemella)
        try:
            sch = leggi_json(p / "scheda_pec.json", {})
            sch["copie"] = [c for c in (sch.get("copie") or []) if c.get("chiave") != k]
            sch["copie"].append(
                {"chiave": k, "cartella_outlook": testo(prop(cart, "FolderPath")),
                 "ricevuta": str(quando)[:19], "mittente": mit})
            scrivi_json(p / "scheda_pec.json", sch)
        except Exception as e:
            err(f"aggiornamento copia in {p.name}: {e}")
        visti_mail[k] = {"cartella": p.name, "oggetto": ogg[:120],
                         "ricevuta": str(quando)[:19], "documenti": 0,
                         "copia_di": firma}
        voce_referto["copia_perche"] = regola_firma
        voce_referto["copia_regola"] = regola_firma
        voce_referto["copia_stesso_lancio"] = firma in FIRME_LANCIO
        voce_referto["copia_di"] = p.name[:70]
        # Come sopra: chiave stabile, non k. Il registro vede una riga sola
        # per comunicazione, anche se la PEC arriva in tre caselle diverse.
        if DB_ATTIVO:
            try:
                # PATCH COPIA: arricchisci il dict con i campi utili.
                # Se c'e' una gemella, eredita la sua scheda completa.
                scheda_copia = {}
                if gemella:
                    try:
                        _p = Path(gemella)
                        _sch_gemella = leggi_json(_p / "scheda_pec.json", {})
                        if _sch_gemella:
                            scheda_copia = _sch_gemella.copy()
                    except Exception:
                        pass

                dict_copia = {
                    "chiave": "copia:" + firma,
                    "ricevuta": str(quando)[:19],
                    "oggetto": ogg,
                    "oggetto_reale": reale,
                    "mittente": mit,
                    "mittente_pec": mit,
                    "tipo_pec": tipo,
                    "fase_ctu": fase,
                    "riferimenti": rif,
                    "novita": [],
                }
                # Arricchisci con i campi della scheda gemella (se disponibile)
                for _k in ["testo", "documenti", "ricorrente", "sede", "grado",
                           "giudice", "data_udienza"]:
                    _v = scheda_copia.get(_k)
                    if _v and not dict_copia.get(_k):
                        dict_copia[_k] = _v
                # Arricchisci anche con i campi che sono gia' nel contesto
                if rif.get("sede"): dict_copia["sede"] = rif["sede"]
                if rif.get("grado"): dict_copia["grado"] = rif["grado"]

                DB.scrivi_copia_archiviata(dict_copia)
            except Exception as e:
                err(f"scrittura db della copia: {e}")
        salva_archivio({"visti": visti_mail, "copie": COPIE,
                        "aggiornato": datetime.now().isoformat(timespec="seconds")})
        log(f"    = copia di {p.name}  ({testo(prop(cart, 'Name'))})  {reale[:44]}")
        return "copia"

    if solo_avviso:
        # Da dove viene. Un messaggio senza allegati, mandato da una casella
        # dell'Istituto a una casella dell'Istituto, non e' una notifica di
        # controparte: e' posta interna. Il controllo di e_inoltro_interno qui
        # non arriva, perche' pretende anche il prefisso FWD:/I:/R: nell'oggetto,
        # e questi non ce l'hanno. Sui 4.030 avvisi in archivio sono 3.290.
        # Non si buttano: vanno in un elenco a parte, cosi' restano ritrovabili.
        _dest = testo(prop(m, "To", "")) + " " + testo(prop(m, "CC", ""))
        _dentro = lambda x: bool(re.search(
            r"@(?:inps\.it|postacert\.inps\.gov\.it|itinps\.onmicrosoft\.com"
            r"|portalecomunicazione\.inps\.it)", testo(x), re.I))
        _esterno_fra_i_dest = bool(re.search(
            r"@(?!inps\.it|postacert\.inps\.gov\.it|itinps\.onmicrosoft\.com)"
            r"[a-z0-9.-]+\.[a-z]{2,}", _dest, re.I))
        # Non basta "interno e senza allegati". Una comunicazione che porta un
        # numero di ruolo o di pratica parla di un fascicolo, e va instradata
        # anche se arriva da un collega e senza documenti: liquidazioni,
        # richieste di integrazione spese, sospensioni. Sull'archivio di oggi
        # sono 900 comunicazioni vere, su fascicoli che esistono, finite fra
        # gli accantonati per questo motivo.
        _porta_riferimento = bool((rif or {}).get("rg") or (rif or {}).get("sisco")
                                  or (rif or {}).get("sentenza"))
        # una ricevuta o una risposta automatica resta rumore anche col ruolo
        _e_ricevuta = bool(RE_RICEVUTA.match(testo(ogg))
                           or re.match(r"^\s*risposta automatica\s*:", testo(ogg), re.I))

        # CHI MANDA NON PUO' ANNULLARE COSA C'E' SCRITTO.
        # "Notificazione ai sensi della legge n. 53 del 1994" e' un criterio
        # forte: la PEC era gia' stata riconosciuta come valida. Poi questo
        # filtro la metteva da parte perche' il mittente e' un collega e non
        # c'erano allegati - e un inoltro di notifica, che e' l'unico modo in
        # cui certe comunicazioni arrivano, spariva dalla lavorazione.
        # Un segnale forte viene dall'OGGETTO, cioe' da cosa la PEC dice di
        # essere: quello decide. Il mittente conta solo quando l'oggetto non
        # ha detto niente di deciso.
        _entrata_forte = testo(motivo).startswith("forte:")

        if SCARTA_INTERNE and _dentro(mit) and not _esterno_fra_i_dest \
           and not _entrata_forte \
           and (not _porta_riferimento or _e_ricevuta):
            ACCANTONATI.append({
                "chiave": k, "ricevuta": str(quando)[:19], "mittente": mit,
                "destinatari": _dest[:200], "oggetto": reale, "tipo": tipo,
                "riferimenti": rif, "motivo": "provenienza interna senza allegati",
                "cartella_outlook": testo(prop(cart, "Name"))})
            visti_mail[k] = {"cartella": None, "oggetto": ogg[:120],
                             "ricevuta": str(quando)[:19], "documenti": 0,
                             "tipo": tipo, "accantonato": "interno"}
            segna_copia(firma, "__avviso__")
            voce_referto["accantonato_interno"] = True
            log(f"    - da parte, interno  ({reale[:46]})")
            return "accantonato"

        # anche senza allegati il modulo di cancelleria da' ruolo, parte e
        # provvedimento: l'avviso non e' piu' una riga muta
        voce_referto["modulo_tribunale"] = bool(scheda_trib)
        uff_a = sede_e_grado([corpo_avviso, reale], mit)
        if uff_a.get("sede"):
            rif["sede"] = uff_a["sede"]
        if uff_a.get("grado"):
            rif["grado"] = uff_a["grado"]
        _u = trova_udienza(f"{corpo_avviso} {reale}")
        _ud_avviso = _u.isoformat() if hasattr(_u, "isoformat") else (_u or None)
        fasc_a, come_a, _ = trova_fascicolo(rif, trib.get("parte"), _u)
        novita_avviso = leggi_atto(
            [trib.get("descrizione"), trib.get("oggetto"), corpo_avviso, reale],
            quando, trib, None, rif)
        AVVISI.append({
            "chiave": k, "ricevuta": str(quando)[:19], "mittente": mit,
            "oggetto": reale, "tipo": tipo, "riferimenti": rif,
            "ricorrente": trib.get("parte") or nome_da_pratica(reale),
            "da_recuperare": e_avviso_da_recuperare(reale, mit, allegati, rif),
            "giudice": trib.get("giudice"),
            "giudice_seguito": giudice_seguito(trib.get("giudice"), CFG,
                                               uff_a.get("grado"), uff_a.get("sede")),
            "provvedimento": trib.get("provvedimento"),
            "numero_provvedimento": trib.get("numero_provvedimento"),
            "comunicazione_tribunale": scheda_trib,
            "sede": uff_a.get("sede"), "grado": uff_a.get("grado"),
            "protocollo": portale.get("protocollo"),
            "ricevuta_inps": portale.get("ricevuta_inps"),
            "mittente_reale": portale.get("mittente_reale"),
            "in_coda_a": fasc_a.name if fasc_a else None,
            "agganciata_per": come_a,
            "cartella_outlook": testo(prop(cart, "Name")),
            # mancava: senza questa nessun avviso poteva risultare completo.
            # Qui siamo prima del calcolo di `udienza` (fatto piu' avanti sul
            # testo dei documenti, che l'avviso non ha), quindi si cerca nel
            # corpo e nell'oggetto, che e' tutto quello che c'e'.
            "data_udienza": _ud_avviso,
            "termine_caricamento": (quando + timedelta(days=30)).strftime("%Y-%m-%d")
                                    if quando else None,
        })
        if DB_ATTIVO:
            try:
                DB.scrivi_scheda({
                    "chiave": k,
                    "ricevuta": str(quando)[:19],
                    "oggetto": reale,
                    "mittente": mit,
                    "tipo_pec": tipo,
                    "riferimenti": rif,
                    "novita": novita_avviso,
                    "notifica_valida": None,
                    "casella_notifica": None,
                    "termine_caricamento":
                        (quando + timedelta(days=30)).strftime("%Y-%m-%d")
                        if quando else None,
                    "ricorrente": trib.get("parte") or nome_da_pratica(reale),
                    "data_udienza": _ud_avviso,
                    "comunicazione_tribunale": scheda_trib,
                    "da_recuperare": e_avviso_da_recuperare(reale, mit, allegati, rif),
                }, tipo_esito="avviso")
            except Exception as e:
                err(f"scrittura db dell'avviso: {e}")
        voce_referto["avviso_instradato"] = bool(fasc_a)
        if fasc_a:
            try:
                LANCIO.setdefault("code", []).append(
                    {"fascicolo": str(fasc_a), "box": None, "chiave": k})
                aggiungi_comunicazione(fasc_a, {
                    "ricevuta": str(quando)[:19], "tipo": tipo, "oggetto": reale,
                    "mittente": mit, "riferimenti": rif, "agganciata_per": come_a,
                    "cartella": None, "documenti": [],
                    "giudice": trib.get("giudice"),
                    "provvedimento": trib.get("provvedimento"),
                    "numero_provvedimento": trib.get("numero_provvedimento"),
                    # anche senza allegati il modulo di cancelleria dice dei
                    # fatti: rinvio, provvedimento, esito. Si leggono con la
                    # stessa funzione, cosi' la procedura e' una sola per
                    # tutte le PEC.
                    "novita": novita_avviso,
                    "da_confermare": sum(1 for x in novita_avviso
                                         if x.get("certezza") != "alta"),
                    "nota": "segnalazione senza documenti: i file stanno nell'intranet",
                    "chiave": k, "registrata_il": datetime.now().isoformat(timespec="seconds")})
                log(f"    > segnalata a {fasc_a.name}  (senza documenti)  {come_a}")
                log(f"      dal modulo: {fatti_in_chiaro(novita_avviso)}")
            except Exception as e:
                err(f"registro comunicazioni di {fasc_a.name}: {e}")
        visti_mail[k] = {"cartella": None, "oggetto": ogg[:120],
                         "ricevuta": str(quando)[:19], "documenti": 0,
                         "tipo": tipo, "solo_avviso": True}
        segna_copia(firma, "__avviso__")
        LANCIO.setdefault("chiavi", []).append(k)
        return "avviso"

    # La firma si prenota PRIMA di creare la cartella. Registrandola in fondo
    # (riga ~2772) restava scoperto tutto il tratto in mezzo: se il messaggio
    # gemello passava di qui prima che il primo avesse finito, o se il primo
    # usciva da un ramo che non arrivava in fondo, COPIE non lo sapeva ancora
    # e nasceva una seconda cartella "(2)". Sulle 1.958 schede in IN ARRIVO
    # questo ha prodotto 474 cartelle in eccesso su 245 comunicazioni.
    prossima_cartella.chiave_corrente = k   # per riconoscere la cartella gia' sua
    dest = prossima_cartella(dest_radice, quando, tipo, rif, reale)
    segna_copia(firma, str(dest))
    dest.mkdir(parents=True, exist_ok=True)
    LANCIO.setdefault("cartelle_transito", []).append(str(dest))
    LANCIO.setdefault("chiavi", []).append(k)
    esito = {"file": [], "note": [], "doppioni": 0, "testo": ""}
    visti_file = set()

    tmp = dest / "_tmp"
    trovata_busta = False
    try:
        for j in range(1, m.Attachments.Count + 1):
            a = m.Attachments.Item(j)
            nome = pulisci_nome(prop(a, "FileName"))
            if nome.lower() in SALTA and not e_la_busta(nome) \
               and Path(nome.lower()).name != "daticert.xml":
                continue
            file_dentro, corpo, nota = contenuto_allegato(a, tmp)
            if nota:
                esito["note"].append(nota)
                continue
            if corpo:
                esito["testo"] += "\n" + corpo
            if e_la_busta(nome):
                trovata_busta = True
            for n2, d2 in file_dentro:
                if n2.lower().endswith(".eml") or e_la_busta(n2):
                    esplodi_eml(d2, dest, visti_file, esito)
                else:
                    scarta_involucri(d2, n2, dest, visti_file, esito)
    except Exception as e:
        esito["note"].append(f"lettura allegati: {e}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    if not trovata_busta:
        esito["note"].append("nessuna busta postacert: trattata come mail ordinaria")
        esito["testo"] += "\n" + testo(prop(m, "Body"))

    # sugli avvisi del Portale la casella PEC compare nel corpo dell'avviso,
    # non nei destinatari Outlook: qui il corpo si legge, sono pochi messaggi
    if corpo_avviso and "nessun postacert.eml" not in " ".join(esito["note"]):
        esito["testo"] = (esito["testo"] + "\n" + corpo_avviso).strip()
    notifica = casella_di_notifica(m, (esito["testo"] or "") + " " + corpo_avviso,
                                   CASELLE, esito.get("daticert"), esito.get("daticert_tutti"))
    voce_referto["notifica"] = notifica["notifica_valida"]
    voce_referto["fonte_notifica"] = notifica.get("fonte")
    voce_referto["casella_notifica"] = notifica.get("casella")
    voce_referto["inoltro"] = "inoltrata internamente" in (notifica.get("difetto") or "")
    voce_referto["daticert"] = bool(esito.get("daticert"))
    voce_referto["n_attestazioni"] = len(esito.get("daticert_tutti") or [])

    if not esito["file"]:
        # allegati dichiarati ma nulla di estraibile: la cartella sarebbe vuota.
        # Va in elenco come avviso, con il motivo, e la cartella si rimuove.
        motivo_vuoto = ("nessun documento estratto da: " + ", ".join(allegati[:6])
                        if allegati else "nessun allegato in casella")
        try:
            for x in dest.iterdir():
                x.unlink()
            dest.rmdir()
        except Exception:
            pass
        nt = casella_di_notifica(m, testo(prop(m, "Body")), CASELLE,
                                 esito.get("daticert"), esito.get("daticert_tutti"))
        voce_referto["notifica"] = nt["notifica_valida"]
        voce_referto["fonte_notifica"] = nt.get("fonte")
        voce_referto["casella_notifica"] = nt.get("casella")
        voce_referto["inoltro"] = "inoltrata internamente" in (nt.get("difetto") or "")
        voce_referto["daticert"] = bool(esito.get("daticert"))
        voce_referto["n_attestazioni"] = len(esito.get("daticert_tutti") or [])
        fasc0, come0, creata0 = trova_fascicolo(rif, None, None)
        novita_vuota = leggi_atto(
            [trib.get("descrizione"), trib.get("oggetto"), reale,
             esito.get("testo"), testo(prop(m, "Body"))],
            quando, trib, esito.get("xml_atto"), rif)
        if fasc0:
            try:
                LANCIO.setdefault("code", []).append(
                    {"fascicolo": str(fasc0), "box": None, "chiave": k})
                aggiungi_comunicazione(fasc0, {
                    "ricevuta": str(quando)[:19], "tipo": tipo, "oggetto": reale,
                    "mittente": mit, "riferimenti": rif, "agganciata_per": come0,
                    "cartella": None, "documenti": [],
                    "novita": novita_vuota,
                    "da_confermare": sum(1 for x in novita_vuota
                                         if x.get("certezza") != "alta"),
                    "nota": "segnalazione senza documenti: " + motivo_vuoto,
                    "chiave": k, "registrata_il": datetime.now().isoformat(timespec="seconds")})
                log(f"    > segnalata a {fasc0.name}  (senza documenti)  {come0}")
                log(f"      dal testo: {fatti_in_chiaro(novita_vuota)}")
            except Exception as e:
                err(f"registro comunicazioni di {fasc0.name}: {e}")
        AVVISI.append({
            "chiave": k, "ricevuta": str(quando)[:19], "mittente": mit,
            "oggetto": reale, "tipo": tipo, "riferimenti": rif,
            "ricorrente": (trib.get("parte") or nome_da_pratica(reale + " " + " ".join(allegati))
                           or nome_da_notifica_53(reale)),
            "atto_notificato": dati_da_notifica_53(reale)[1],
            "da_recuperare": e_avviso_da_recuperare(reale, mit, allegati, rif),
            "giudice": trib.get("giudice"),
            "provvedimento": trib.get("provvedimento"),
            "numero_provvedimento": trib.get("numero_provvedimento"),
            "comunicazione_tribunale": scheda_trib,
            "cartella_outlook": testo(prop(cart, "Name")),
            "in_coda_a": fasc0.name if fasc0 else None,
            "motivo": motivo_vuoto,
            "note": esito["note"],
            "termine_caricamento": (quando + timedelta(days=30)).strftime("%Y-%m-%d")
                                    if quando else None,
        })
        if DB_ATTIVO:
            try:
                DB.scrivi_scheda({
                    "chiave": k,
                    "ricevuta": str(quando)[:19],
                    "oggetto": reale,
                    "mittente": mit,
                    "tipo_pec": tipo,
                    "riferimenti": rif,
                    "novita": novita_vuota,
                    "notifica_valida": nt["notifica_valida"],
                    "casella_notifica": nt["casella"],
                    "termine_caricamento":
                        (quando + timedelta(days=30)).strftime("%Y-%m-%d")
                        if quando else None,
                    "ricorrente": (trib.get("parte")
                                   or nome_da_pratica(reale + " " + " ".join(allegati))
                                   or nome_da_notifica_53(reale)),
                    "data_udienza": None,
                    "comunicazione_tribunale": scheda_trib,
                    "da_recuperare": e_avviso_da_recuperare(reale, mit, allegati, rif),
                }, tipo_esito="avviso")
            except Exception as e:
                err(f"scrittura db dell'avviso: {e}")
        visti_mail[k] = {"cartella": None, "oggetto": ogg[:120],
                         "ricevuta": str(quando)[:19], "documenti": 0,
                         "tipo": tipo, "solo_avviso": True, "motivo": motivo_vuoto}
        segna_copia(firma, "__avviso__")
        log(f"    ~ senza documenti  ({motivo_vuoto[:52]})  {reale[:40]}")
        return "avviso"

    scheda = {
        "chiave": k,
        "cartella_outlook": testo(prop(cart, "FolderPath")),
        "ricevuta": str(quando)[:19],
        "mittente": mit,
        "mittente_pec": (esito.get("daticert") or {}).get("mittente"),
        "destinatari": testo(prop(m, "To")),
        "destinatari_certificati": (esito.get("daticert") or {}).get("destinatari_certificati"),
        "oggetto_certificato": (esito.get("daticert") or {}).get("oggetto"),
        "oggetto": ogg,
        "oggetto_reale": reale,
        "tipo_pec": tipo,
        "fase_ctu": fase,
        "riferimenti": rif,
        "protocollo": portale.get("protocollo"),
        "ricevuta_inps": portale.get("ricevuta_inps"),
        "mittente_reale": portale.get("mittente_reale") or (esito.get("daticert") or {}).get("mittente"),
        "casella_notifica": notifica["casella"],
        "notifica_valida": notifica["notifica_valida"],
        "difetto_notifica": notifica["difetto"],
        "notifica_fonte": notifica.get("fonte"),
        "daticert": esito.get("daticert"),
        "motivo_selezione": motivo,
        "documenti": esito["file"],
        "doppioni_scartati": esito["doppioni"],
        "note": esito["note"],
        "stato": ("da_analizzare" if esito["file"]
                  else ("solo_avviso" if not allegati else "da_verificare")),
        "termine_caricamento": (quando + timedelta(days=30)).strftime("%Y-%m-%d") if quando else None,
        "elaborata_il": datetime.now().isoformat(timespec="seconds"),
    }
    if esito["testo"].strip():
        (dest / "testo_pec.txt").write_text(esito["testo"].strip(), encoding="utf-8")

    # --- nome definitivo: RICORRENTE UD ggmmaa, come le cartelle dell'ufficio ---
    # ricorsivo: in un inoltro (FWD) i PDF stanno dentro il messaggio originale
    testo_doc, nota_pdf = testo_dei_documenti(dest, ricorsivo=True)
    if nota_pdf:
        esito["note"].append(nota_pdf)
    tutto = (testo_doc or "") + "\n" + (esito["testo"] or "") + "\n" + reale

    # IL MODULO STA DOVE CAPITA: nel corpo esterno, o dentro la busta, o in
    # entrambi a meta'. Si leggono tutti e due e si uniscono campo per campo,
    # con il testo della busta per primo: li' il modulo arriva come l'ha
    # scritto la cancelleria, mentre nel corpo esterno ci passa sopra
    # l'inoltro dell'avvocato, che ne mangia i campi.
    # Prima valeva "il primo che risponde", e un corpo esterno con mezzo
    # modulo bloccava la lettura di quello interno.
    trib = modulo_di_cancelleria(esito["testo"] or "", corpo_avviso)
    scheda_trib = {k: v for k, v in trib.items() if v not in (None, "", False)} or None
    if trib.get("rg") and not rif.get("rg"):
        rif["rg"] = trib["rg"]
    xmlatto = esito.get("xml_atto") or {}
    # Terza fonte: il modulo dentro <Contenuto> dell'XML della busta. Su una
    # comunicazione telematica il corpo Outlook e' vuoto e i PDF sono scansioni:
    # li' dentro c'e' l'unico testo strutturato. leggi_xml_atto ora lo legge,
    # ma il risultato finiva in xmlatto e non arrivava alla scheda: parte e
    # giudice restavano vuoti anche quando erano stati trovati.
    if xmlatto:
        trib = dict(trib or {})
        for k, v in xmlatto.items():
            if v not in (None, "", False) and not trib.get(k):
                trib[k] = v
        scheda_trib = {k: v for k, v in trib.items() if v not in (None, "", False)} or None
        if trib.get("rg") and not rif.get("rg"):
            rif["rg"] = trib["rg"]
    # L'identificativo del gestore non cambia: se questa comunicazione l'abbiamo
    # gia' presa, l'inoltro che gira mesi dopo e' la stessa cosa. Serve la busta
    # aperta, quindi il controllo sta qui e non prima.
    # L'identificativo del gestore e' unico per messaggio: due messaggi che
    # lo portano uguale SONO la stessa PEC. Prima il controllo si faceva solo
    # sugli oggetti con prefisso di inoltro, per paura di scartare l'originale
    # insieme alla copia - ma il primo che passa registra e viene tenuto, il
    # secondo trova l'identificativo gia' preso. La paura era infondata, e il
    # prezzo era alto: le copie con oggetto identico e senza prefisso non
    # venivano riconosciute. Nella prova dal 2026-08-01 la firma ne prendeva
    # 20 su 22, e quelle che sfuggivano erano tutte di questa forma: oggetto
    # di rito, riferimenti vuoti, arrivate a minuti diversi.
    ids = identificativo_originale(esito, testo(corpo))
    for idm in ids:
        if IDENTIFICATIVI.get(idm.lower()):
            visti_mail[k] = {"cartella": None, "oggetto": ogg[:120],
                             "ricevuta": str(quando)[:19], "documenti": 0,
                             "copia_di": firma, "inoltro_di": idm}
            voce_referto["copia_perche"] = "identificativo del gestore"
            voce_referto["copia_di"] = testo(IDENTIFICATIVI.get(idm.lower()))[:70]
            log(f"    = stessa PEC gia' presa (id gestore)  ({reale[:44]})")
            return "copia"
    # Si registra SUBITO, come la firma: se lo si fa alla fine, il gemello che
    # passa nel frattempo non lo trova e nasce una seconda cartella.
    for idm in ids:
        IDENTIFICATIVI.setdefault(idm.lower(), "in lavorazione")

    ufficio = sede_e_grado(
        [corpo_avviso, esito.get("testo") or "", reale,
         (trib or {}).get("sede") or "", xmlatto.get("sede") or "",
         # i nomi dei file: la sigla (CT) sta spesso solo li'
         " ".join(str(d.get("nome") or "") for d in (esito.get("file") or []))],
        (esito.get("daticert") or {}).get("mittente") or mit)
    if ufficio.get("sede"):
        rif["sede"] = ufficio["sede"]
    if ufficio.get("grado"):
        rif["grado"] = ufficio["grado"]
    if xmlatto.get("rg") and not rif.get("rg"):
        rif["rg"] = xmlatto["rg"]
    voce_referto["modulo_tribunale"] = bool(scheda_trib)

    # IL GIUDICE PUO' STARE SOLO DENTRO IL DOCUMENTO.
    # Finora si cercava nel modulo di cancelleria e nell'XML della busta.
    # Ma un inoltro di un collega non ha ne' l'uno ne' l'altro: l'oggetto
    # dice "citazione testi", il mittente e' @inps.it, e il nome del giudice
    # compare solo dentro il PDF allegato. Quelle PEC passavano tutti i
    # controlli e finivano fra i ricorsi da aprire anche quando erano di
    # Bologna o di Gurrieri.
    # Il testo dei documenti e' gia' stato letto qui sopra (testo_doc): si
    # usa, invece di lasciarlo inutilizzato. Le due funzioni che servono
    # esistono gia' e sono prudenti: giudice_dal_testo pretende il titolo
    # "dott."/"dr.", giudice_dalla_firma accetta solo nomi gia' noti fra i
    # giudici della sezione.
    giudice_visto = ((trib or {}).get("giudice") or xmlatto.get("giudice")
                     or giudice_dal_testo(tutto) or giudice_dalla_firma(tutto))
    if giudice_visto and not ((trib or {}).get("giudice") or xmlatto.get("giudice")):
        voce_referto["giudice_dal_documento"] = giudice_visto
    voce_referto["giudice"] = giudice_visto
    voce_referto["giudice_altro"] = (giudice_seguito(
        giudice_visto, CFG,
        ufficio.get("grado"), ufficio.get("sede")) is False)
    voce_referto["xml_atto"] = bool(xmlatto)
    voce_referto["riferimenti"] = rif

    # Le PEC che non ti competono non vanno mescolate alle tue. Sta QUI e non
    # piu' avanti: da qui in poi ci sono tre uscite - accodata, unita, avviso -
    # e chi esce di la' non verrebbe mai spostato.
    accantonata_in = None
    if voce_referto.get("giudice_altro"):
        # perche' non viene spostata: e' l'unico modo per saperlo
        voce_referto["perche_non_spostata"] = (
            "accantonamento spento" if not ACCANTONA
            else "la cartella non esiste piu' (nessun documento estratto)"
            if not dest.exists() else None)
    # Le esecuzioni le lavora un altro ufficio: si mettono da parte, non si
    # buttano. Restano in ESECUZIONI e nella scheda Scartate del registro.
    if ESECUZIONI and dest.exists() and not voce_referto.get("giudice_altro") \
       and e_esecuzione(reale, corpo_avviso, (trib or {}).get("oggetto"),
                        (trib or {}).get("descrizione"),
                        # anche il TESTO dei documenti: un inoltro di un
                        # collega non ha la sezione nell'oggetto, ma l'atto
                        # dentro dice "Tribunale di Siracusa, sezione
                        # esecuzioni mobiliari". Solo le prime pagine: la
                        # sezione sta in testa, e cercarla in fondo trova le
                        # citazioni di altri procedimenti.
                        (testo_doc or "")[:4000],
                        (esito.get("testo") or "")[:4000]):
        try:
            nuova_radice = dest_radice / "ESECUZIONI"
            nuova_radice.mkdir(parents=True, exist_ok=True)
            spostata = nuova_radice / dest.name
            if not spostata.exists():
                dest.rename(spostata)
                ct = LANCIO.get("cartelle_transito") or []
                if str(dest) in ct:
                    ct[ct.index(str(dest))] = str(spostata)
                dest = spostata
                accantonata_in = "ESECUZIONI"
                voce_referto["esecuzione"] = True
                log("    \u2192 ESECUZIONI")
        except Exception as e:
            esito["note"].append(f"non spostata in ESECUZIONI: {e}")

    if ACCANTONA and voce_referto.get("giudice_altro") and dest.exists():
        try:
            nuova_radice = dest_radice / "ALTRI GIUDICI"
            nuova_radice.mkdir(parents=True, exist_ok=True)
            spostata = nuova_radice / dest.name
            if not spostata.exists():
                dest.rename(spostata)
                ct = LANCIO.get("cartelle_transito") or []
                if str(dest) in ct:
                    ct[ct.index(str(dest))] = str(spostata)
                dest = spostata
                accantonata_in = "ALTRI GIUDICI"
                voce_referto["arrivata_ad_accantonare"] = True
                log("    \u2192 ALTRI GIUDICI")
        except Exception as e:
            esito["note"].append(f"non spostata in ALTRI GIUDICI: {e}")
    # prima il modulo, che la scrive in forma fissa; poi la ricerca generica.
    # La differenza fra le due va SCRITTA: la prima e' una data dichiarata
    # dalla cancelleria, la seconda e' solo la prima data vicina alla parola
    # "udienza" in tutto il testo, allegati compresi. Serve a ritrovare la
    # cartella, non a fissare un'udienza - e senza la targa il registro non
    # poteva distinguerle, cosi' una data presa da un modulo ricorrente
    # finiva proposta come rinvio su decine di fascicoli.
    udienza = udienza_dal_modulo(trib.get("descrizione"), trib.get("oggetto"))
    udienza_fonte = "modulo di cancelleria" if udienza else None
    if not udienza:
        udienza = trova_udienza(tutto)
        if udienza:
            udienza_fonte = "prima data accanto alla parola 'udienza' nel testo"
    # La data NUOVA, quella del rinvio: si cerca nel modulo E nel testo degli
    # allegati - il verbale arriva sempre come PDF - e si tiene solo se sta
    # dopo la ricezione della PEC. E' il dato che mancava: senza, il registro
    # vedeva la comunicazione ma non sapeva che portava un'altra udienza, e la
    # data andava battuta a mano due volte, qui e in SISCO.
    _rif_data = quando.date() if hasattr(quando, "date") else None
    udienza_nuova, udienza_frase = udienza_annunciata(
        [trib.get("descrizione"), trib.get("oggetto"), tutto], _rif_data)
    if not udienza and (esito.get("xml_atto") or {}).get("udienza"):
        try:
            a2, m2, g2 = esito["xml_atto"]["udienza"].split("-")
            udienza = date(int(a2), int(m2), int(g2))
            udienza_fonte = "XML dell'atto"
        except Exception:
            pass
    # Le fonti del nominativo, in ordine di attendibilita'. Le prime due sono
    # strutturate; le altre sono deduzioni, e la piu' debole - il nome del
    # file - viene per ultima, solo se tutte le altre hanno taciuto.
    # Tre giri, come la conta di magazzino: oggetto, testo dell'atto, nomi dei
    # file. Il nome vale se due giri lo leggono uguale; se no resta vuoto. Il
    # primo giro da solo sbagliava (una societa' al posto della persona, un
    # nome preso da un'altra causa). La parte data dal tribunale e' un dato,
    # non una lettura: vale da sola.
    letture_nome = [
        ("oggetto", nome_da_pratica(reale) or nome_da_notifica_53(reale) or nome_dalla_causa(reale)),
        ("atto", trova_ricorrente(testo_doc or "", reale)),
        ("file", nome_da_pratica(" ".join(allegati)) or nome_da_nomefile(allegati)),
    ]
    letture_nome = [(g, n) for g, n in letture_nome if n and not perche_non_e_un_nome(n)]
    ricorrente = (trib.get("parte")
                  or (esito.get("xml_atto") or {}).get("parte")
                  or nome_a_tre_giri(letture_nome))
    # Il giro dell'atto leggeva solo le prime 4 pagine (e le ultime 2) di 6
    # PDF: quando il nome sta piu' avanti - la procura in fondo, il decreto
    # unito al ricorso, un inoltro con molti allegati - l'atto taceva e il nome
    # restava vuoto. Se i tre giri non bastano, l'atto si rilegge per intero e
    # si rifa' la conta: vale sempre la regola dei due giri uguali.
    if not ricorrente:
        pieno, _n = testo_dei_documenti(dest, max_pagine=MAX_PAGINE_ATTI, max_file=12, ricorsivo=True)
        if pieno and len(pieno) > len(testo_doc or ""):
            n_pieno = trova_ricorrente(pieno, reale)
            if n_pieno and not perche_non_e_un_nome(n_pieno):
                letture_nome = [(g, n) for g, n in letture_nome if g != "atto"] + [("atto", n_pieno)]
                ricorrente = nome_a_tre_giri(letture_nome)
                if ricorrente:
                    log(f"    + nominativo letto nel documento intero: {testo(ricorrente)[:40]}")
                    testo_doc = pieno
    if not ricorrente and letture_nome:
        log("    ~ nominativo non confermato: " + " / ".join(f"{g}: {testo(n)[:30]}" for g, n in letture_nome))
    motivo_nome = perche_non_e_un_nome(ricorrente) if ricorrente else None
    if motivo_nome:
        log(f"    ~ nominativo scartato: '{testo(ricorrente)[:40]}' \u2014 {motivo_nome}")
        ricorrente = None
    da_nomefile = False    # il nome file da solo non basta piu': serve un secondo giro

    if not rif.get("rg"):
        r2 = riferimenti(tutto[:20000])
        for kk, vv in r2.items():
            rif.setdefault(kk, vv)

    # --- il fascicolo esiste gia'? allora questa e' una comunicazione, non un
    #     fascicolo nuovo: va in coda dentro la sua cartella COMUNICAZIONI ---
    fasc, come, creata = trova_fascicolo(rif, ricorrente, udienza)
    if not fasc and tipo != "prima_istanza":
        if come and "sede diversa" in str(come):
            motivo_no_coda = come
        elif not rif:
            motivo_no_coda = "nessun riferimento nell'oggetto ne' nei documenti"
        elif not (FASC["per_rg"] or FASC["cartella_rg"]):
            motivo_no_coda = "instradamento spento: manca l'indice del registro"
        else:
            rgn = re.sub(r"[^0-9]", "", str((rif or {}).get("rg") or ""))
            scn = _compatta((rif or {}).get("sisco") or "")
            voce_reg = (FASC["per_sisco"].get(scn) if scn else None) or \
                       (FASC["per_rg"].get(rgn)
                        if rgn and rgn not in FASC["rg_ambigui"] else None)
            if voce_reg and ((rif or {}).get("grado") == "appello"):
                motivo_no_coda = (f"fascicolo {voce_reg.get('ricorrente') or ''} in appello: "
                                  "riconosciuto dal ruolo ma senza cartella "
                                  "(in appello il SISCO non c'e')")
            elif voce_reg and not voce_reg.get("udienze"):
                motivo_no_coda = (f"fascicolo {voce_reg.get('ricorrente') or ''} in registro, "
                                  "ma senza udienze: il nome della cartella non e' componibile")
            elif voce_reg and not FASC["puo_creare"]:
                motivo_no_coda = (f"fascicolo {voce_reg.get('ricorrente') or ''} in registro, "
                                  "cartella assente e creazione disattivata "
                                  "(puo_creare_fascicoli)")
            elif voce_reg:
                motivo_no_coda = (f"fascicolo {voce_reg.get('ricorrente') or ''} in registro, "
                                  "ma la cartella non e' stata creata: controlla i permessi "
                                  "su archivio_fascicoli")
            else:
                riff = ", ".join(f"{k} {v}" for k, v in (rif or {}).items())
                motivo_no_coda = f"nessun fascicolo in registro per {riff}"
        esito["note"].append("non instradata: " + motivo_no_coda)
    else:
        motivo_no_coda = None

    # I fatti che gli allegati portano: si leggono PRIMA di accodare, cosi' la
    # comunicazione entra nel fascicolo gia' con dentro cosa dice.
    novita = leggi_atto([trib.get("descrizione"), trib.get("oggetto"),
                         (esito.get("xml_atto") or {}).get("oggetto"), tutto],
                        quando, trib, esito.get("xml_atto"), rif,
                        [d["nome"] for d in esito["file"]], oggetto=reale)
    voce_referto["novita"] = len(novita)
    voce_referto["fatti"] = fatti_in_chiaro(novita)

    if fasc and tipo != "prima_istanza":
        try:
            coda = fasc / "COMUNICAZIONI"
            coda.mkdir(exist_ok=True)
            etichetta = pulisci_nome(
                " - ".join(x for x in [str(quando)[:10],
                                       (tipo or "").replace("_", " "),
                                       (fase or ""), reale[:40]] if x), 100)
            box = coda / etichetta
            n = 2
            while box.exists():
                box = coda / f"{etichetta} ({n})"
                n += 1
            box.mkdir()
            for x in sorted(dest.iterdir()):
                if x.is_file():
                    shutil.move(str(x), str(box / x.name))
            shutil.rmtree(dest, ignore_errors=True)

            LANCIO.setdefault("code", []).append(
                {"fascicolo": str(fasc), "box": str(box), "chiave": k})
            aggiungi_comunicazione(fasc, {
                "ricevuta": str(quando)[:19], "tipo": tipo, "fase_ctu": fase,
                "oggetto": reale, "mittente": mit, "riferimenti": rif,
                "agganciata_per": come, "cartella": box.name,
                "documenti": [d["nome"] for d in esito["file"]],
                # la data annunciata dal verbale: il registro la propone con
                # un click, non la si ribatte a mano
                "data_udienza": udienza.isoformat() if udienza else None,
                "udienza_fonte": udienza_fonte,
                "udienza_indicata": udienza_nuova.isoformat() if udienza_nuova else None,
                "udienza_come": udienza_frase,
                # TUTTI i fatti degli allegati, non solo il rinvio: sentenza,
                # esito, spese, CTU, termini, riferimenti nuovi. E' questo che
                # rende la comunicazione un aggiornamento del fascicolo
                # invece di un documento parcheggiato dentro.
                "novita": novita,
                "da_confermare": sum(1 for x in novita if x.get("certezza") != "alta"),
                "notifica_valida": notifica["notifica_valida"],
                "chiave": k, "registrata_il": datetime.now().isoformat(timespec="seconds"),
            })

            visti_mail[k] = {"cartella": str(box), "oggetto": ogg[:120],
                             "ricevuta": str(quando)[:19],
                             "documenti": len(esito["file"]), "in_coda_a": fasc.name}
            segna_copia(firma, str(box))
            for idm in ids:
                IDENTIFICATIVI.setdefault(idm.lower(), str(box))
            IN_CODA.append({"fascicolo": fasc.name, "cartella": box.name,
                            "fascicolo_creato": creata,
                            "ricevuta": str(quando)[:19], "tipo": tipo, "fase_ctu": fase,
                            "oggetto": reale, "riferimenti": rif, "agganciata_per": come,
                            "documenti": len(esito["file"]),
                            "novita": novita,
                            "termine_caricamento": None})
            salva_archivio({"visti": visti_mail, "copie": COPIE, "avvisi": AVVISI, "accantonati": ACCANTONATI,
                                   "aggiornato": datetime.now().isoformat(timespec="seconds")})
            log(f"    > in coda a {fasc.name}"
                + ("  [CARTELLA CREATA]" if creata else "")
                + f"  ({len(esito['file'])} doc)  {come}")
            log(f"      dagli allegati: {fatti_in_chiaro(novita)}")
            voce_referto["in_coda"] = True
            if DB_ATTIVO:
                try:
                    DB.scrivi_scheda({
                        "chiave": k,
                        "ricevuta": str(quando)[:19],
                        "oggetto": ogg,
                        "oggetto_reale": reale,
                        "mittente": mit,
                        "mittente_pec": (esito.get("daticert") or {}).get("mittente"),
                        "tipo_pec": tipo,
                        "fase_ctu": fase,
                        "riferimenti": rif,
                        "novita": novita,
                        "notifica_valida": notifica["notifica_valida"],
                        "casella_notifica": notifica["casella"],
                        "termine_caricamento":
                            (quando + timedelta(days=30)).strftime("%Y-%m-%d")
                            if quando else None,
                        "sede": ufficio.get("sede"),
                        "grado": ufficio.get("grado"),
                        "giudice": trib.get("giudice"),
                        "ricorrente": trib.get("parte") or ricorrente,
                        "data_udienza": udienza.isoformat() if udienza else None,
                        "comunicazione_tribunale": scheda_trib,
                        "cartella": str(fasc),
                        "accantonata_in": accantonata_in,
                    }, tipo_esito="coda")
                except Exception as e:
                    err(f"scrittura db della coda: {e}")
            return "coda"
        except Exception as e:
            esito["note"].append(f"accodamento non riuscito: {e}")

    giudice = (trib or {}).get("giudice") or xmlatto.get("giudice")
    mio = giudice_seguito(giudice, CFG, ufficio.get("grado"), ufficio.get("sede"))

    if mio is False:
        esito["note"].append(f"giudice {giudice_canonico(giudice)}: non fra quelli seguiti, "
                             "il fascicolo lo tiene un altro operatore")
        voce_referto["giudice_altro"] = True

    # --- scarto tardivo: comunicazioni interne che viaggiano in risposta a PEC ---
    # Si riconoscono da cosa NON hanno all'ultimo livello: nessun XML dell'atto,
    # nessun modulo di cancelleria, nessuna attestazione del gestore. Il tipo
    # resta indeterminato perche' il contenuto non e' strutturato.
    if SCARTA_INTERNE and tipo in ("da_classificare", "da_leggere") \
       and not scheda_trib and not xmlatto and not esito.get("daticert"):
        mit_interno = bool(re.search(r"@inps\.it|@postacert\.inps\.gov\.it", testo(mit), re.I))
        risposte = len(re.findall(r"\b(?:risposta|re|rif|i|fw|fwd|inoltro)\s*:", ogg, re.I))
        motivo_int = ("nessun contenuto strutturato (n\u00e9 XML dell'atto, n\u00e9 modulo di "
                      "cancelleria, n\u00e9 attestazione del gestore)")
        if mit_interno or risposte >= 2:
            motivo_int += ": comunicazione interna in risposta a una PEC"
            shutil.rmtree(dest, ignore_errors=True)
            ct = LANCIO.get("cartelle_transito") or []
            if str(dest) in ct:
                ct.remove(str(dest))
            visti_mail[k] = {"cartella": None, "oggetto": ogg[:120],
                             "ricevuta": str(quando)[:19], "documenti": 0,
                             "scartata": motivo_int}
            voce_referto["scartata_tardi"] = motivo_int
            log(f"    x interna, scartata  ({reale[:46]})")
            return "interna"

    # LA REGOLA: vedi e_pec_da_fascicolo. Una PEC senza nessuno dei segni di
    # un atto e' posta di servizio, e aprirle un fascicolo affoga quelli veri.
    da_fascicolo, perche_no = e_pec_da_fascicolo(
        esito, {**(esito.get("xml_atto") or {}), **rif,
                "mittente_pec": esito.get("mittente_pec"),
                "da_modulo": bool(trib), "ufficio_giudiziario": (trib or {}).get("ufficio")},
        allegati, mit, ogg, corpo)
    if not da_fascicolo:
        tipo = "da_classificare"
        esito["note"].append("non e' una PEC da fascicolo: " + perche_no)
        voce_referto["declassata"] = True
        voce_referto["perche_no"] = perche_no

    definitivo, completo = nome_definitivo(ricorrente, udienza, rif, quando, tipo, reale)
    if definitivo:
        nuova = dest.parent / definitivo
        if nuova.exists() and nuova.resolve() != dest.resolve():
            # il fascicolo esiste gia': i documenti si aggiungono, non si duplicano
            try:
                aggiunti = unisci_cartelle(dest, nuova, esito)
                LANCIO.setdefault("unioni", []).append(
                    {"cartella": str(nuova), "file": aggiunti, "chiave": k})
                vecchia = leggi_json(nuova / "scheda_pec.json", {})
                vecchia["copie"] = [c for c in (vecchia.get("copie") or [])
                                    if c.get("chiave") != k]
                vecchia["copie"].append({
                    "chiave": k, "cartella_outlook": testo(prop(cart, "FolderPath")),
                    "ricevuta": str(quando)[:19], "mittente": mit,
                    "oggetto": reale, "tipo": tipo,
                    "documenti_aggiunti": aggiunti})
                vecchia["documenti"] = (vecchia.get("documenti") or []) + \
                    [{"nome": n2, "da": "PEC del " + str(quando)[:10]} for n2 in aggiunti]
                for kk, vv in (rif or {}).items():
                    vecchia.setdefault("riferimenti", {}).setdefault(kk, vv)
                if esito["testo"].strip():
                    p_t = nuova / "testo_pec.txt"
                    prec = p_t.read_text(encoding="utf-8") if p_t.exists() else ""
                    p_t.write_text((prec + "\n\n" + "-"*60 + "\nPEC del " +
                                    str(quando)[:16] + "\n" + esito["testo"].strip()).strip(),
                                   encoding="utf-8")
                scrivi_json(nuova / "scheda_pec.json", vecchia)
                shutil.rmtree(dest, ignore_errors=True)
                for i, s in enumerate(LAVORATE):
                    if s.get("cartella") == str(nuova):
                        LAVORATE[i] = vecchia
                        break
                # anche l'unita' va marcata, altrimenti l'accantonamento non
                # risulta da nessuna parte: e sono la maggioranza
                if accantonata_in:
                    vecchia["accantonata_in"] = accantonata_in
                    scrivi_json(nuova / "scheda_pec.json", vecchia)
                visti_mail[k] = {"cartella": nuova.name, "oggetto": ogg[:120],
                                 "ricevuta": str(quando)[:19], "accantonata_in": accantonata_in,
                                 "documenti": len(aggiunti), "unita_a": nuova.name}
                segna_copia(firma, str(nuova))
                for idm in ids:
                    IDENTIFICATIVI.setdefault(idm.lower(), str(nuova))
                salva_archivio({"visti": visti_mail, "copie": COPIE, "avvisi": AVVISI, "accantonati": ACCANTONATI,
                                       "aggiornato": datetime.now().isoformat(timespec="seconds")})
                log(f"    + {len(aggiunti)} documenti aggiunti a {nuova.name}")
                voce_referto["unita"] = True
                return "unita"
            except Exception as e:
                esito["note"].append(f"fusione non riuscita: {e}")
        try:
            if nuova.resolve() != dest.resolve():
                dest.rename(nuova)
                ct = LANCIO.get("cartelle_transito") or []
                if str(dest) in ct:
                    ct[ct.index(str(dest))] = str(nuova)
                dest = nuova
        except Exception as e:
            esito["note"].append(f"rinomina non riuscita: {e}")
    # DA ESAMINARE si decide solo qui: prima il tipo non e' definitivo.
    # Chi e' gia' in ALTRI GIUDICI non si tocca.
    if ACCANTONA and not accantonata_in and dest.exists() \
       and tipo in ("da_classificare", "da_leggere") and not rif:
        try:
            nr = dest_radice / "DA ESAMINARE"
            nr.mkdir(parents=True, exist_ok=True)
            sp = nr / dest.name
            if not sp.exists():
                dest.rename(sp)
                ct = LANCIO.get("cartelle_transito") or []
                if str(dest) in ct:
                    ct[ct.index(str(dest))] = str(sp)
                dest = sp
                accantonata_in = "DA ESAMINARE"
                log("    \u2192 DA ESAMINARE")
        except Exception as e:
            esito["note"].append(f"non spostata in DA ESAMINARE: {e}")

    if definitivo and not completo:
        esito["note"].append("nome parziale: manca la data di udienza, "
                             "la cartella non e' agganciabile dal registro finche' non si aggiunge")
    if not definitivo:
        esito["note"].append("nome provvisorio: ricorrente non ricavabile dai documenti")

    scheda["comunicazione_tribunale"] = scheda_trib
    scheda["xml_atto"] = esito.get("xml_atto")
    scheda["sede"] = ufficio.get("sede")
    scheda["grado"] = ufficio.get("grado")
    if not scheda.get("giudice"):
        scheda["giudice"] = (esito.get("xml_atto") or {}).get("giudice")
    scheda["giudice"] = giudice
    scheda["giudice_seguito"] = mio
    scheda["provvedimento"] = (trib or {}).get("provvedimento")
    scheda["numero_provvedimento"] = (trib or {}).get("numero_provvedimento")
    scheda["inps_ricorrente"] = (trib or {}).get("inps_ricorrente")
    scheda["ricorrente"] = ricorrente
    scheda["ricorrente_letture"] = ({g: n for g, n in letture_nome} if (letture_nome and not ricorrente) else None)
    scheda["atto_notificato"] = dati_da_notifica_53(reale)[1]
    scheda["data_udienza"] = udienza.isoformat() if udienza else None
    scheda["udienza_indicata"] = udienza_nuova.isoformat() if udienza_nuova else None
    scheda["udienza_come"] = udienza_frase
    # i fatti restano scritti anche quando la PEC apre un fascicolo nuovo:
    # il ricorso dice l'udienza, il decreto il giudice, la sentenza l'esito
    scheda["novita"] = novita
    scheda["da_confermare"] = sum(1 for x in novita if x.get("certezza") != "alta")
    if e_avviso_da_recuperare(reale, mit, allegati, rif):
        # non e' un difetto di lettura: e' un atto che sta altrove
        scheda["da_recuperare"] = True
        voce_referto["da_recuperare"] = True
    if da_nomefile:
        # dedotto, non letto: va segnalato, altrimenti un nome sbagliato
        # entra in archivio senza che nessuno possa accorgersene
        scheda["nome_dedotto_da"] = "nome del file"
    scheda["motivo_non_instradata"] = motivo_no_coda
    scheda["nome_convenzione"] = definitivo if completo else None
    scheda["nome_parziale"] = None if completo else definitivo
    scheda["riferimenti"] = rif
    scheda["firma_comunicazione"] = firma
    scheda["copie"] = []
    scheda["cartella"] = str(dest)
    scheda["accantonata_in"] = accantonata_in
    scrivi_json(dest / "scheda_pec.json", scheda)
    LAVORATE.append(scheda)
    segna_copia(firma, str(dest))

    visti_mail[k] = {"cartella": dest.name, "oggetto": ogg[:120],
                     "ricevuta": str(quando)[:19], "documenti": len(esito["file"]),
                     "tipo": tipo, "notifica_valida": notifica["notifica_valida"]}
    salva_archivio({"visti": visti_mail, "copie": COPIE, "avvisi": AVVISI, "accantonati": ACCANTONATI,
                           "aggiornato": datetime.now().isoformat(timespec="seconds")})
    marchio = ("" if notifica["notifica_valida"] else
               "  [DIFETTO DI NOTIFICA]" if notifica["notifica_valida"] is False else
               "  [casella ignota]")
    if DB_ATTIVO:
        try:
            DB.scrivi_scheda({
                "chiave": k,
                "ricevuta": str(quando)[:19],
                "oggetto": ogg,
                "oggetto_reale": reale,
                "mittente": mit,
                "mittente_pec": (esito.get("daticert") or {}).get("mittente"),
                "tipo_pec": tipo,
                "fase_ctu": fase,
                "riferimenti": rif,
                "novita": novita,
                "notifica_valida": notifica["notifica_valida"],
                "casella_notifica": notifica["casella"],
                "termine_caricamento":
                    (quando + timedelta(days=30)).strftime("%Y-%m-%d")
                    if quando else None,
                "sede": ufficio.get("sede"),
                "grado": ufficio.get("grado"),
                "giudice": giudice,
                "ricorrente": ricorrente,
                "data_udienza": udienza.isoformat() if udienza else None,
                "comunicazione_tribunale": scheda_trib,
                "cartella": str(dest),
                "accantonata_in": accantonata_in,
            }, tipo_esito="nuovo")
        except Exception as e:
            err(f"scrittura db del fascicolo: {e}")

    log(f"    -> {dest.name}  ({len(esito['file'])} documenti){marchio}  {reale[:52]}")
    return True


# --------------------------------------------------------------------------
def dataIt(d):
    try:
        return datetime.strptime(str(d)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return str(d or "")
def scrivi_indice(radice, lavorate):
    """Un elenco unico ordinato per urgenza: senza questo, 142 cartelle
    sono solo 142 domande su quale aprire per prima."""
    if not lavorate and not AVVISI:
        return
    oggi_s = datetime.now().strftime("%Y-%m-%d")

    def giorni(t):
        try:
            return (datetime.strptime(t, "%Y-%m-%d") - datetime.now()).days
        except Exception:
            return 999

    righe = []
    for x in lavorate:
        righe.append({**x,
            "tipo": x.get("tipo_pec") or x.get("tipo"),
            "fase_ctu": x.get("fase_ctu"),
            "oggetto": x.get("oggetto_reale") or x.get("oggetto"),
            "documenti": len(x.get("documenti", [])), "avviso": False,
            "ricorrente": x.get("ricorrente"), "data_udienza": x.get("data_udienza"),
            "notifica_valida": x.get("notifica_valida"),
            "casella_notifica": x.get("casella_notifica"),
            "protocollo": x.get("protocollo"),
            "ricevuta_inps": x.get("ricevuta_inps"),
            "mittente_reale": x.get("mittente_reale"),
            "giudice": x.get("giudice"),
            "provvedimento": x.get("provvedimento"),
            "numero_provvedimento": x.get("numero_provvedimento"),
            "motivo_non_instradata": x.get("motivo_non_instradata")})
    for a in AVVISI:
        righe.append({**a, "documenti": 0, "cartella": "", "avviso": True,
                      "data_udienza": a.get("data_udienza")})
    righe.sort(key=lambda x: (x.get("termine_caricamento") or "9999",
                              x.get("ricevuta") or ""))

    def rif_breve(r):
        r = r or {}
        for e, c in (("SISCO", "sisco"), ("RG", "rg"), ("sent", "sentenza"),
                     ("com", "comunicazione"), ("spese", "spese_legali")):
            if r.get(c):
                return f"{e} {r[c]}"
        return ""

    q = lambda v: '"' + str(v if v is not None else "").replace('"', '""') + '"'

    # Le PEC accantonate non entrano nell'elenco di lavoro: sono di altri
    # o non classificabili. Restano contate a parte, e sul disco ci sono.
    accantonate = [x for x in righe if x.get("accantonata_in")]
    righe = [x for x in righe if x not in accantonate]

    # --- elenco leggibile ---
    L = []
    add = L.append
    scadute = [x for x in righe if giorni(x.get("termine_caricamento")) < 0]
    urgenti = [x for x in righe if 0 <= giorni(x.get("termine_caricamento")) <= 7]
    add("INDICE DELLE PEC LAVORATE - " + datetime.now().strftime("%d/%m/%Y %H:%M"))
    add("=" * 100)
    add(f"cartelle con documenti ..... {len(lavorate)}")
    # Le esecuzioni: si contano dalle righe dell'indice, non dal referto.
    # Questa riga era copiata dalla funzione del referto, dove 'referto'
    # esiste; qui no, e l'eccezione veniva ingoiata dal try che avvolge tutta
    # la scrittura: l'indice usciva mutilato e nessuno lo sapeva.
    nEsec = sum(1 for x in righe
                if str(x.get("accantonata_in") or "").upper() == "ESECUZIONI")
    if nEsec:
        add(f"esecuzioni messe da parte .. {nEsec}"
            "   sezione esecuzioni: le lavora un altro ufficio")
    if ACCANTONATI:
        # Dichiarato, non silenzioso: uno scarto che non si vede non lo
        # scopre nessuno finche' non si ricontano i messaggi a mano.
        add(f"messi da parte (interni) ... {len(ACCANTONATI)}"
            "   senza allegati, da casella INPS a casella INPS")
        add("                              sono in archivio_pec.json -> accantonati")
    viziate = sum(1 for x in righe if x.get("notifica_valida") is False)
    if viziate:
        add(f"con difetto di notifica ... {viziate}  (notificate a casella non deputata)")
    add(f"in coda a fascicoli aperti . {len(IN_CODA)}" +
        ("" if (FASC["per_chiave"] or FASC["per_sisco"])
         else "   <-- instradamento SPENTO: mancano archivio_fascicoli e base_json"))
    if FASC["per_chiave"] or FASC["per_sisco"]:
        add(f"  (indice: {len(FASC['per_chiave'])} cartelle, {FASC['quante']} fascicoli in registro)")
    add(f"  (una cartella si crea solo se contiene almeno un documento apribile)")
    add(f"avvisi senza documenti ..... {len(AVVISI)}  (nessuna cartella: i documenti stanno nel Portale)")
    # LE FRASI DELLA CANCELLERIA CHE NON SAPPIAMO LEGGERE. Vanno dette: una
    # descrizione muta e' una sentenza, un rinvio o una CTU che il fascicolo
    # non impara, e il silenzio e' indistinguibile da "non c'era niente".
    if DESCRIZIONI_MUTE:
        tot_mute = sum(DESCRIZIONI_MUTE.values())
        add("")
        add(f"DESCRIZIONI DEL MODULO NON LETTE: {tot_mute} "
            f"in {len(DESCRIZIONI_MUTE)} forme diverse")
        add("  Il modulo diceva qualcosa e non siamo riusciti a capire cosa.")
        add("  I numeri sono sostituiti da N: e' la FORMA della frase che conta.")
        for forma, n in sorted(DESCRIZIONI_MUTE.items(), key=lambda kv: -kv[1])[:20]:
            add(f"    {n:5d}  {forma}")
        if len(DESCRIZIONI_MUTE) > 20:
            add(f"    ... e altre {len(DESCRIZIONI_MUTE)-20} forme")
    if accantonate:
        add(f"accantonate ................ {len(accantonate)}")
        alt = sum(1 for x in accantonate if x.get("accantonata_in") == "ALTRI GIUDICI")
        if alt:
            add(f"   {alt} in ALTRI GIUDICI: le tiene un altro operatore")
        if len(accantonate) - alt:
            add(f"   {len(accantonate)-alt} in DA ESAMINARE: nessun riferimento riconosciuto")
    add(f"oltre i 30 giorni .......... {len(scadute)}")
    add(f"in scadenza entro 7 giorni . {len(urgenti)}")
    add("")
    add("Ordinate per termine di caricamento: le prime sono quelle che scadono prima.")
    add("")

    con_doc = [x for x in righe if x.get("cartella")]
    senza = [x for x in righe if not x.get("cartella")]

    add("=" * 100)
    add(f"DA LAVORARE: {len(con_doc)} cartelle con documenti")
    add("=" * 100)
    add(f"{'TERMINE':<12}{'GG':>5}  {'RICORRENTE':<26}{'UDIENZA':<12}{'DOC':>4}  {'TIPO':<26}RIFERIMENTO")
    add("-" * 110)
    for x in con_doc:
        g = giorni(x.get("termine_caricamento"))
        segno = "SCADUTO" if g < 0 else (f"{g} gg" if g < 999 else "-")
        ud = x.get("data_udienza")
        add(f"{(x.get('termine_caricamento') or '-'):<12}{segno:>5}  "
            f"{(x.get('ricorrente') or '(da leggere)'):<26}"
            f"{(dataIt(ud) if ud else '-'):<12}{(x['documenti'] or ''):>4}  "
            f"{(x.get('tipo') or '-'):<26}{rif_breve(x.get('riferimenti'))}")
        if x.get("notifica_valida") is False:
            add(f"{'':<49}! DIFETTO DI NOTIFICA: {x.get('casella_notifica') or ''}")
        if x.get("motivo_non_instradata"):
            add(f"{'':<49}? non instradata: {x['motivo_non_instradata']}")
        if x.get("cartella"):
            add(f"{'':<49}\\ {Path(x['cartella']).name}")
        else:
            add(f"{'':<49}\\ {x.get('motivo') or 'avviso del Portale: i documenti stanno nell intranet'}")
    def causa(m):
        """Il motivo contiene i numeri della singola pratica: per capire dove
        intervenire serve la causa, non l'elenco dei casi."""
        m = testo(m)
        if "nessun fascicolo in registro" in m:
            return ("ruolo non presente nel registro",
                    "il fascicolo non e' mai stato caricato, o ha un numero diverso")
        if "senza udienze" in m:
            return ("fascicolo nel registro ma senza udienze",
                    "sono i fascicoli storici: senza data non si compone il nome cartella")
        if "sede diversa" in m:
            return ("stesso ruolo ma in un'altra sede",
                    "aggancio evitato: Siracusa e Catania usano la stessa numerazione")
        if "creazione disattivata" in m:
            return ("cartella assente e creazione disattivata",
                    "metti puo_creare_fascicoli a true per farle nascere")
        if "nessun riferimento" in m:
            return ("nessun riferimento leggibile",
                    "ne' oggetto ne' documenti portano un numero utilizzabile")
        return ("altro", "")

    per_causa, dettagli = {}, {}
    for x in righe:
        if x.get("motivo_non_instradata"):
            k, spiega = causa(x["motivo_non_instradata"])
            per_causa[k] = per_causa.get(k, 0) + 1
            dettagli.setdefault(k, [spiega, []])
            if len(dettagli[k][1]) < 5:
                dettagli[k][1].append(testo(x["motivo_non_instradata"])[:88])
    if per_causa:
        tot_ni = sum(per_causa.values())
        add("")
        add("=" * 100)
        add(f"COMUNICAZIONI NON INSTRADATE: {tot_ni}")
        add("=" * 100)
        add("Hanno creato una cartella in IN ARRIVO invece di entrare nel fascicolo.")
        add("")
        for k, n in sorted(per_causa.items(), key=lambda y: -y[1]):
            add(f"  {n:5d}  ({round(100*n/tot_ni):2d}%)  {k}")
            if dettagli[k][0]:
                add(f"{'':<15}{dettagli[k][0]}")
            for d in dettagli[k][1][:3]:
                add(f"{'':<15}es. {d}")
            add("")
    if IN_CODA:
        add("")
        add("=" * 100)
        add(f"IN CODA A FASCICOLI GIA' APERTI: {len(IN_CODA)} comunicazioni")
        add("=" * 100)
        add("Non hanno creato una cartella nuova: i documenti stanno dentro il fascicolo,")
        add("nella sottocartella COMUNICAZIONI, e sono registrati in comunicazioni.json")
        add("")
        nuovi = sum(1 for x in IN_CODA if x.get("fascicolo_creato"))
        if nuovi:
            add(f"di cui {nuovi} in cartelle create ora dall'ingestione (contrassegnate con *)")
            add("")
        add(f"{'FASCICOLO':<34}{'DATA':<12}{'TIPO':<26}{'DOC':>4}  AGGANCIATA PER")
        add("-" * 110)
        for x in sorted(IN_CODA, key=lambda y: (y["fascicolo"], y["ricevuta"])):
            add(("*" if x.get("fascicolo_creato") else " ") +
                f"{x['fascicolo'][:32]:<33}{x['ricevuta'][:10]:<12}"
                f"{((x.get('tipo') or '') + (' / ' + x['fase_ctu'] if x.get('fase_ctu') else ''))[:25]:<26}"
                f"{x['documenti']:>4}  {x.get('agganciata_per') or ''}")
    add("")
    add("=" * 100)
    add(f"SOLO SEGNALAZIONE: {len(senza)} comunicazioni senza documenti in casella")
    add("=" * 100)
    add("Non c'e' nulla da aprire: servono per il termine dei 30 giorni e per i riferimenti.")
    add("")
    add(f"{'TERMINE':<12}{'GG':>5}  {'TIPO':<26}{'RIFERIMENTO':<22}OGGETTO")
    add("-" * 100)
    for x in senza:
        g = giorni(x.get("termine_caricamento"))
        segno = "SCADUTO" if g < 0 else (f"{g} gg" if g < 999 else "-")
        add(f"{(x.get('termine_caricamento') or '-'):<12}{segno:>5}  "
            f"{(x.get('tipo') or '-'):<26}"
            f"{rif_breve(x.get('riferimenti')):<22}{testo(x.get('oggetto'))[:44]}")
        if x.get("motivo") and "avviso del Portale" not in x["motivo"]:
            add(f"{'':<45}\\ {x['motivo'][:70]}")
    (radice / "_INDICE.txt").write_text("\n".join(L), encoding="utf-8")

    # --- i ruoli che il registro non conosce: e' una lista di lavoro ---
    ignoti = {}
    for x in righe:
        m = testo(x.get("motivo_non_instradata"))
        if "nessun fascicolo in registro" not in m:
            continue
        r = x.get("riferimenti") or {}
        chiave = r.get("rg") or r.get("sisco") or r.get("comunicazione")
        if not chiave:
            continue
        v = ignoti.setdefault(chiave, {
            "rg": r.get("rg"), "sisco": r.get("sisco"),
            "sede": r.get("sede"), "grado": r.get("grado"),
            "ricorrente": x.get("ricorrente"), "giudice": x.get("giudice"),
            "provvedimento": x.get("provvedimento"),
            "numero_provvedimento": x.get("numero_provvedimento"),
            "pec": 0, "prima": x.get("ricevuta"), "ultima": x.get("ricevuta"),
            "oggetti": []})
        v["pec"] += 1
        v["ricorrente"] = v["ricorrente"] or x.get("ricorrente")
        v["giudice"] = v["giudice"] or x.get("giudice")
        if x.get("ricevuta"):
            v["prima"] = min(v["prima"] or x["ricevuta"], x["ricevuta"])
            v["ultima"] = max(v["ultima"] or x["ricevuta"], x["ricevuta"])
        # l'oggetto certificato porta dentro tutta la ricevuta: qui serve
        # solo la prima riga utile, non il testo del messaggio
        og = testo(x.get("oggetto"))
        og = re.split(r"Messaggio di posta certificata|Certified email message|"
                      r"Il giorno \d{2}/\d{2}/\d{4}", og)[0]
        og = re.sub(r"^(?:POSTA CERTIFICATA:|I:|R:|FWD?:|RE:|RIF:)\s*", "", og.strip(), flags=re.I)
        og = re.sub(r"\s+", " ", og).strip()[:70]
        if og and og not in v["oggetti"] and len(v["oggetti"]) < 4:
            v["oggetti"].append(og)

    if ignoti:
        # quanti anni hanno: se sono vecchi non sono ricorsi nuovi sfuggiti,
        # sono fascicoli che il registro non ha mai avuto
        anni = {}
        for v in ignoti.values():
            a = (v.get("rg") or "").split("/")[-1]
            if len(a) == 4:
                anni[a] = anni.get(a, 0) + 1
        vecchi = sum(n for a, n in anni.items() if a < str(datetime.now().year))

        # chi ha lo stesso nominativo in archivio: o e' un'altra causa dello
        # stesso ricorrente, o il ruolo e' trascritto male
        per_nome = {}
        for f in FASC.get("per_rg", {}).values():
            n = _compatta(f.get("ricorrente"))[:10]
            if n:
                per_nome.setdefault(n, []).append(f)
        simili = []
        for kk, v in ignoti.items():
            n = _compatta(v.get("ricorrente"))[:10]
            for f in per_nome.get(n, []):
                a = (v.get("rg") or "").split("/")
                c = (f.get("rg") or "").split("/")
                if len(a) == 2 and len(c) == 2 and a[0] == c[0] and a[1] != c[1]:
                    simili.append((v.get("rg"), f.get("rg"), v.get("ricorrente"),
                                   f.get("ricorrente")))
                    break

        F = ["RUOLI CHE IL REGISTRO NON CONOSCE - " + datetime.now().strftime("%d/%m/%Y %H:%M"),
             "=" * 100,
             f"{len(ignoti)} procedimenti con {sum(v['pec'] for v in ignoti.values())} PEC "
             "che non trovano un fascicolo."] + \
            ([f"{vecchi} hanno un ruolo di anni precedenti: non sono ricorsi nuovi",
              "sfuggiti, sono procedimenti che il registro non ha mai avuto.",
              "Per anno di iscrizione: " +
              ", ".join(f"{a}: {n}" for a, n in sorted(anni.items()))]
             if vecchi else []) + \
            [
             "",
             "Ogni riga e' un fascicolo che il tribunale considera vivo e il registro non ha:",
             "o non e' mai stato caricato, o e' registrato con un numero diverso.",
             "Controllali in SISCO e creali dove mancano: da quel momento le loro",
             "comunicazioni entreranno da sole nel fascicolo."] + \
            ([""] + ["=" * 100,
             f"ATTENZIONE: {len(simili)} hanno lo STESSO NUMERO di un fascicolo in archivio,",
             "ma con anno diverso. Uno dei due e' trascritto male: controlla quale.",
             "=" * 100,
             f"{'NELLA PEC':<14}{'IN ARCHIVIO':<14}RICORRENTE"] +
             [f"{testo(a):<14}{testo(b):<14}{testo(c)[:34]:<36}{testo(d)[:26]}"
              for a, b, c, d in simili[:20]]
             if simili else []) + \
            [
             "",
             f"{'RUOLO':<14}{'SEDE':<11}{'PEC':>4}  {'RICORRENTE':<28}{'GIUDICE':<22}PERIODO",
             "-" * 100]
        for k, v in sorted(ignoti.items(), key=lambda y: -y[1]["pec"]):
            # il nome puo' essere lungo: va troncato, altrimenti si mangia
            # la colonna del giudice e le due si attaccano
            nome = testo(v["ricorrente"] or "(non ricavabile)")[:26]
            giud = giudice_canonico(v["giudice"] or "")[:20]
            F.append(f"{(v['rg'] or k):<14}{(v['sede'] or '?'):<11}{v['pec']:>4}  "
                     f"{nome:<28}{giud:<22}"
                     f"{str(v['prima'])[:10]} - {str(v['ultima'])[:10]}")
            if v["provvedimento"]:
                F.append(f"{'':<14}{v['provvedimento']}"
                         + (f" n. {v['numero_provvedimento']}" if v["numero_provvedimento"] else ""))
            for o in v["oggetti"][:2]:
                F.append(f"{'':<14}{o}")
            F.append("")
        (radice / "_RUOLI_SCONOSCIUTI.txt").write_text("\n".join(F), encoding="utf-8")
        try:      # una copia dove --azzera non arriva
            (BASE / "_RUOLI_SCONOSCIUTI.txt").write_text("\n".join(F), encoding="utf-8")
        except Exception:
            pass

        colI = ["RUOLO", "SISCO", "SEDE", "GRADO", "RICORRENTE", "GIUDICE",
                "PROVVEDIMENTO", "N. PROVVEDIMENTO", "PEC RICEVUTE", "DALLA", "ALLA", "OGGETTO"]
        CI = ["\ufeff" + ";".join(q(c) for c in colI)]
        for k, v in sorted(ignoti.items(), key=lambda y: -y[1]["pec"]):
            CI.append(";".join(q(x) for x in [
                v["rg"] or k, v["sisco"], v["sede"], v["grado"], v["ricorrente"], v["giudice"],
                v["provvedimento"], v["numero_provvedimento"], v["pec"],
                str(v["prima"])[:10], str(v["ultima"])[:10],
                " | ".join(v["oggetti"])]))
        (radice / "_ruoli_sconosciuti.csv").write_text("\r\n".join(CI), encoding="utf-8")
        try:
            (BASE / "_ruoli_sconosciuti.csv").write_text("\r\n".join(CI), encoding="utf-8")
        except Exception:
            pass
        log(f"_RUOLI_SCONOSCIUTI: {len(ignoti)} procedimenti senza fascicolo nel registro")
        log(f"  copia anche in {BASE}, che --azzera non tocca")

    # --- versione per Excel ---
    col = ["TERMINE 30 GG", "GIORNI", "RICEVUTA", "RICORRENTE", "UDIENZA", "TIPO", "FASE CTU", "DOCUMENTI",
           "RIFERIMENTO", "SISCO", "RG", "SENTENZA", "COMUNICAZIONE", "SPESE",
           "NOTIFICA VALIDA", "CASELLA NOTIFICA", "PROTOCOLLO", "RICEVUTA INPS",
           "GIUDICE", "PROVVEDIMENTO", "OGGETTO", "MITTENTE", "CARTELLA"]
    C = ["\ufeff" + ";".join(q(c) for c in col)]
    for x in righe:
        r = x.get("riferimenti") or {}
        g = giorni(x.get("termine_caricamento"))
        C.append(";".join(q(v) for v in [
            x.get("termine_caricamento"), (g if g < 999 else ""), x.get("ricevuta"),
            x.get("ricorrente"), x.get("data_udienza"),
            x.get("tipo"), x.get("fase_ctu"), x["documenti"], rif_breve(r),
            r.get("sisco"), r.get("rg"), r.get("sentenza"), r.get("comunicazione"),
            r.get("spese_legali"), x.get("notifica_valida"), x.get("casella_notifica"),
            x.get("protocollo"), x.get("ricevuta_inps"),
            x.get("giudice"), (x.get("provvedimento") or "") +
              (" n. " + x["numero_provvedimento"] if x.get("numero_provvedimento") else ""),
            testo(x.get("oggetto"))[:200],
            x.get("mittente_reale") or x.get("mittente_pec") or x.get("mittente"),
            Path(x["cartella"]).name if x.get("cartella") else "(solo avviso)"]))
    (radice / "_indice.csv").write_text("\r\n".join(C), encoding="utf-8")
    log(f"indice scritto: _INDICE.txt e _indice.csv ({len(righe)} voci)")

def _sblocca(funzione, percorso, _):
    """Su OneDrive capita che i file siano in sola lettura o ancora agganciati
    dalla sincronizzazione: si toglie l'attributo e si riprova."""
    try:
        os.chmod(percorso, stat.S_IWRITE)
        funzione(percorso)
    except Exception:
        pass

def rimuovi(percorso, tentativi=6):
    """Cancella una cartella o un file gestendo i blocchi di OneDrive.
    Se proprio non si riesce, la mette da parte rinominandola: meglio una
    cartella con un nome che dice 'da cancellare' che un errore e basta."""
    p = Path(percorso)
    for n in range(tentativi):
        try:
            if not p.exists():
                return True
            if p.is_dir():
                # prima i file uno a uno, poi le cartelle dal basso:
                # rmtree si ferma al primo file bloccato, cosi' invece si
                # recupera tutto il resto
                for x in sorted(p.rglob("*"), key=lambda y: -len(y.parts)):
                    try:
                        if x.is_file():
                            os.chmod(x, stat.S_IWRITE)
                            x.unlink()
                        elif x.is_dir():
                            x.rmdir()
                    except Exception:
                        pass
                shutil.rmtree(p, onerror=_sblocca)
            else:
                os.chmod(p, stat.S_IWRITE)
                p.unlink()
            if not p.exists():
                return True
        except Exception:
            pass
        time.sleep(0.5 * (n + 1))      # la sincronizzazione molla la presa
    try:
        fuori = p.with_name("_DA CANCELLARE " + p.name)
        n = 2
        while fuori.exists():
            fuori = p.with_name(f"_DA CANCELLARE {p.name} ({n})")
            n += 1
        p.rename(fuori)
        err(f"{p.name}: bloccata da OneDrive, rinominata in '{fuori.name}' "
            "(cancellala a mano quando la sincronizzazione ha finito)")
        return False
    except Exception as e:
        err(f"{p.name}: non rimossa e non rinominabile ({e})")
        return False

# fra la parola "giudice" e il nome ci possono stare parecchie formule:
# "Il Giudice del Lavoro designato dott.ssa Marescalco Anna"
# Nei fascicoli i documenti sono quasi sempre le relate di notifica
# dell'avvocato di controparte: il giudice non lo firma, ma lo NOMINA lui.
#   "innanzi al Giudice delegato Dott. Paolo Marescalco"
#   "dal Giudice del Lavoro , Dott.ssa Giovanna Pedalino ,"
#   "di udienza del 19.03.2025 Dott. Gurrieri"
# Il titolo (dott./dott.ssa) e' il segnale piu' affidabile: quando c'e',
# quello che segue e' un nome di persona.
RE_GIUDICE_TESTO = re.compile(
    r"\b(?:giudice|magistrat[oa])\b"
    r"((?:[\s:,\-]+(?:delegat[oa]|designat[oa]|istruttore|relatore|titolare|"
    r"onorario|unico|lavoro|previdenza|dello|della|del|di|e|sig\.?(?:ra)?))*)"
    r"[\s:,\-]*\(?\s*(?:(?:dott|dr)\.?\s?(?:ssa)?)?[\s.]*"
    r"((?:[A-Z][A-Za-z\u00c0-\u00ff']{2,20}[ ]{0,2}){1,3})", re.I)

# il titolo da solo, ovunque compaia: "Dott.ssa Giovanna Pedalino"
RE_TITOLO_NOME = re.compile(
    r"(?:dott|dr)\.?\s?(?:ssa)?[\s.]*"
    r"((?:[A-Z][A-Za-z\u00c0-\u00ff']{2,20}[ ]{0,2}){1,3})", re.I)
RE_OGGETTO_TESTO = re.compile(
    r"(?:avente\s+ad\s+oggetto|per\s+l['\u2019]?\s*accertamento\s+del(?:la|lo)?|"
    r"\boggetto\b|\bmateria\b)\s*[:\-]?\s*"
    r"([^\n\r.;]{6,70})", re.I)

# in calce ai provvedimenti: "Il Giudice / dott.ssa Anna Marescalco", oppure
# "Firmato da: MARESCALCO ANNA" della firma digitale
RE_FIRMA = re.compile(
    r"(?:il\s+giudice|il\s+magistrato|firmato\s+(?:digitalmente\s+)?da|"
    r"sottoscritto\s+da|f\.?to|firma(?:to)?)"
    # fra la qualifica e il nome: "del lavoro", l'a capo, il titolo
    r"(?:[\s:,\-]+(?:del|dello|della|e|lavoro|previdenza|designat[oa]|"
    r"istruttore|unico|onorario|titolare))*"
    r"[\s:,\-]*\(?[\s]*(?:dott\.?(?:ssa)?|dr\.?(?:ssa)?|avv\.?|sig\.?(?:ra)?)?[\s.]*"
    r"((?:[A-Z][A-Za-z\u00c0-\u00ff']{2,20}[ ]{0,2}){1,3})", re.I)

def giudice_dalla_firma(t):
    """La firma in calce e' del giudice solo nei PROVVEDIMENTI. Nelle relate
    di notifica — che sono la gran parte dei documenti in cartella — quella
    firma e' dell'avvocato che notifica. Quindi qui si accetta solo un nome
    che risulta gia' fra i giudici della sezione."""
    coda = testo(t)[-6000:]
    for m in RE_FIRMA.finditer(coda):
        g = giudice_canonico(m.group(1))
        if scarta_come_nome(g):
            continue
        parole = [p for p in g.split() if len(p) > 2]
        if not parole:
            continue
        v = cognome_prima(" ".join(parole[:2]))
        if any(x in GIUDICI_NOTI for x in v.split()):
            return v
    return None

# "Avv. Rossi", "Avv .Rossi", "avvocato Rossi": il punto puo' stare ovunque
RE_AVVOCATO = re.compile(
    r"avv(?:ocat[oa])?\s*\.?\s*[:\-]?\s*"
    r"((?:[A-Z][A-Za-z\u00c0-\u00ff']{2,20}[ ]{0,2}){1,3})", re.I)

def avvocato_dal_testo(t, scartati=None):
    """Chi firma la relata e' il difensore della parte ricorrente: e' un dato
    che serve, basta non scambiarlo per il giudice. Vale anche un nome trovato
    dopo un titolo ma che non e' un giudice della sezione: e' l'avvocato."""
    tt = testo(t)
    cand = {}
    for v in (scartati or []):
        if not any(x in GIUDICI_NOTI for x in v.split()):
            cand[v] = cand.get(v, 0) + 1
    for m in RE_AVVOCATO.finditer(tt):
        n = giudice_canonico(m.group(1))
        if scarta_come_nome(n):
            continue
        p = [x for x in n.split() if len(x) > 2]
        if not p or (len(p) == 1 and p[0] in NOMI_PROPRI):
            continue
        v = " ".join(p[:2])
        if len(p) >= 2 and p[0] in NOMI_PROPRI and p[1] not in NOMI_PROPRI:
            v = p[1] + " " + p[0]        # "Nino Cortese" -> "CORTESE NINO"
        if any(x in GIUDICI_NOTI for x in v.split()):
            continue          # e' un giudice, non l'avvocato
        cand[v] = cand.get(v, 0) + 1
    if not cand:
        return None
    return max(cand.items(), key=lambda x: x[1])[0]

def cognome_prima(nome, noti=None):
    """"Anna Marescalco" e "Marescalco Anna" sono la stessa persona: il registro
    li tiene per cognome. Se il cognome e' fra quelli conosciuti, si riordina."""
    p = [x for x in giudice_canonico(nome).split() if x]
    if len(p) != 2:
        return " ".join(p)
    conosciuti = set(noti or GIUDICI_NOTI)
    if p[1] in conosciuti and p[0] not in conosciuti:
        return p[1] + " " + p[0]
    # senza un elenco: i nomi propri comuni stanno davanti, il cognome dopo
    NOMI = {"ANNA","MARIA","GIUSEPPE","SALVATORE","ANTONIO","FRANCESCO","LUISA",
            "SERGIO","CARMELO","FILIPPO","GIOVANNI","PAOLO","ROSARIA","LUCIA",
            "ELENA","LAURA","CHIARA","MARCO","ANDREA","ALESSANDRO","DANIELE",
            "SIMONE","SERENA","BARBARA","CRISTINA","VINCENZO","CORRADO"}
    if p[0] in NOMI and p[1] not in NOMI:
        return p[1] + " " + p[0]
    return p[0] + " " + p[1]

GIUDICI_NOTI = set()

def scarta_come_nome(g):
    """Tutto quello che dopo un titolo NON e' un nome di persona: formule di
    rito, verbi, citta', termini delle firme digitali, preposizioni."""
    g = testo(g).strip()
    if len(g.replace(" ", "")) < 4:
        return True
    if re.search(
        r"\b(LAVORO|ISTRUTTORE|DESIGNAT|RELATORE|UNICO|ONORARIO|SEZIONE|"
        r"TRIBUNALE|CORTE|COLLEGIO|PRESIDENTE|FISSA|RITENUT|LETTO|VISTO|"
        r"DISPONE|ORDINA|INVITA|RIGETTA|ACCOGLIE|UDIENZA|RICORSO|ISTANZA|"
        r"SIRACUSA|CATANIA|NOTO|AVOLA|PALERMO|MESSINA|ROMA|MILANO|"
        r"COMPARIRE|PARTI|CAUSA|PROCEDIMENTO|CANCELLIER|SEGRETARI|"
        r"DIGITALMENTE|DOCUMENTO|COPIA|CONFORME|ORIGINALE|ILL|SIG|SPETT|"
        r"MEDICO|PERITO|CTU|CONSULENTE|FUNZIONARIO|DIRETTORE)\b", g):
        return True
    # radici, perche' compaiono flesse: "dichiarante", "utilizzato"
    if re.search(
        r"\b(?:DICHIARANT|UTILIZZAT|TRAMITE|EMESS|SERIAL|CERTIFICAT|TITOLAR|"
        r"SOTTOSCRITT|RILASCIAT|VALIDIT|SCADENZ|CHIAVE|ALGORITM|NOTIFICANT|"
        r"RICHIEDENT|INTERESSAT|DELEGAT|INCARICAT)", g):
        return True
    # frammenti di frase: preposizioni e parole comuni lette come cognome
    COMUNI = {"ATI", "CON", "PER", "TRA", "FRA", "DEL", "DAL", "NEL", "SUL", "COL",
              "GIUSTA", "RISERVA", "SEDE", "ATTO", "ANNO", "MESE", "GIORNO",
              "PARTE", "FINE", "SENSI", "ORDINE", "DATA", "CUI", "CHE", "NON"}
    if all(p in COMUNI for p in g.split() if p):
        return True
    if g.split() and g.split()[0] in COMUNI:
        return True
    return False

def giudice_dal_testo(t):
    """Il giudice si riconosce SOLO se accompagnato dal titolo. Senza, dopo
    la parola "giudice" viene quasi sempre un verbo: "il Giudice voglia
    fissare", "chiede che il Giudice ordini", "il Giudice adito". Sono
    formule di rito, non nomi."""
    tt = testo(t)          # tutto, non solo l'inizio: i documenti sono molti
    for m in RE_GIUDICE_TESTO.finditer(tt):
        if not re.search(r"(?:dott|dr)\.?\s?(?:ssa)?[\s.]*$", m.group(0)[:m.end()-m.start()
                          - len(m.group(2))], re.I):
            continue       # manca il titolo: e' una formula, non una persona
        g = giudice_canonico(m.group(2))
        parole = [p for p in g.split() if len(p) > 2]
        if not parole or scarta_come_nome(g) or len(parole[0]) < 3:
            continue
        v = cognome_prima(correggi_su_noti(" ".join(parole[:2])))
        if giudice_plausibile(v):
            return v
        SCARTATI.append(v)

    # il titolo vicino a un riferimento processuale: probabile magistrato
    for m in RE_TITOLO_NOME.finditer(tt):
        contorno = tt[max(0, m.start() - 120): m.end() + 60]
        if not re.search(r"udienza|r\.?\s*g\.?|procedimento|tribunale|sezione\s+lavoro",
                         contorno, re.I):
            continue
        if re.search(r"\bavv\.?\b|difensor|procurator|ricorrent|resistent",
                     contorno[:120], re.I):
            continue          # e' l'avvocato, non il giudice
        g = giudice_canonico(m.group(1))
        parole = [p for p in g.split() if len(p) > 2]
        if not parole or scarta_come_nome(g) or len(parole[0]) < 3:
            continue
        # un cognome solo va bene: "Dott. Gurrieri" e' comunque il giudice
        v = cognome_prima(correggi_su_noti(" ".join(parole[:2])))
        if giudice_plausibile(v):
            return v
        SCARTATI.append(v)
    return None

def _somiglia(a, b):
    """Due cognomi che differiscono per una lettera sono la stessa persona:
    'PAEDALINO' e' PEDALINO con un refuso di battitura."""
    if abs(len(a) - len(b)) > 1:
        return False
    if a == b:
        return True
    lungo, corto = (a, b) if len(a) >= len(b) else (b, a)
    for i in range(len(lungo)):
        if lungo[:i] + lungo[i+1:] == corto:      # una lettera in piu'
            return True
    if len(a) == len(b):
        return sum(1 for x, y in zip(a, b) if x != y) == 1   # una diversa
    return False

def correggi_su_noti(v):
    """Se un cognome somiglia a uno dei giudici della sezione, e' quello:
    i refusi vengono dalla lettura del PDF, non dall'atto."""
    p = [x for x in testo(v).split() if x]
    for i, x in enumerate(p):
        if x in GIUDICI_NOTI:
            return v
    for i, x in enumerate(p):
        for noto in GIUDICI_NOTI:
            if len(x) > 4 and _somiglia(x, noto):
                p[i] = noto
                return " ".join(p)
    return v

def giudice_plausibile(v):
    """I giudici della sezione lavoro sono un insieme CHIUSO: quelli in
    anagrafica. Un nome che non e' fra quelli non e' un magistrato nuovo,
    e' l'avvocato che ha scritto o notificato l'atto. Se un giudice arriva
    davvero, lo si aggiunge in anagrafica dal registro e da li' in poi
    viene riconosciuto."""
    p = [x for x in testo(v).split() if x]
    if not p:
        return False
    if not GIUDICI_NOTI:
        # senza anagrafica non posso decidere: scarto solo i nomi di battesimo
        return not (len(p) == 1 and p[0] in NOMI_PROPRI)
    return any(x in GIUDICI_NOTI for x in p)

NOMI_PROPRI = {"ANNA","MARIA","GIUSEPPE","GIUSEPPA","SALVATORE","ANTONIO","FRANCESCO",
    "LUISA","SERGIO","CARMELO","FILIPPO","GIOVANNI","GIOVANNA","PAOLO","PAOLA",
    "ROSARIA","LUCIA","LUCA","ELENA","LAURA","CHIARA","MARCO","ANDREA","ALESSANDRO",
    "DANIELE","SIMONE","SERENA","BARBARA","CRISTINA","VINCENZO","CORRADO","MADDALENA",
    "FEDERICO","MAURIZIO","IVANO","MARIO","CINZIA","GIANNICOLA","ROBERTO","SEBASTIANO",
    "NINO","DANIELA","LUCIA","FORTUNATA","DAVIDE","MARCO","CARMELA","ROSSANA","ALESSANDRA",
    "GRAZIA","SERENELLA","MARIANNA","LILY","BARBARA","VINCENZA","CONCETTA","AGATA",
    "SALVATRICE","ANGELA","ROSA","TERESA","CATERINA","GAETANO","ORAZIO","SANTO"}

def oggetto_dal_testo(t):
    m = RE_OGGETTO_TESTO.search(testo(t))
    if not m:
        return None
    v = re.sub(r"\s+", " ", m.group(1)).strip(" .:-\u2013")
    return v if len(v) >= 6 else None

# Codici SISCO verificati su schermate reali. I menu della pagina numerano
# le opzioni in ordine alfabetico, non i codici: vanno confermati uno per uno.
CODICI_MATERIA = {"8": "CONTENZIOSO CONTRIBUTIVO", "19": "INVALIDITA' CIVILE", "22": "INDEBITO", "1": "ASSEGNO INVALIDITA'", "17": "VARIE", "2": "PRESTAZIONI PENSIONISTICHE", "3": "PRESTAZIONI - AMMORTIZZATORI SOCIALI", "14": "FONDI SPECIALI"}
CODICI_ARGOMENTO = {"67": "CONTRIBUTI ARTIGIANI", "183": "ASS.MENSILE DI ASSISTENZA", "184": "PENS. INABILITA'", "185": "IND.ACCOMPAGNAMENTO", "186": "IND.MENSILE FREQUENZA", "241": "HANDICAP EX LEGGE N. 104/92", "234": "INDEBITI DA CONTRIBUTI", "1": "ASS. ORD. INVALIDITA' A.G.O.", "137": "ART. 700", "143": "VARIE", "17": "ASSEGNO/PENS. SOCIALE", "19": "RICOSTITUZIONE PENSIONE", "21": "LIQUIDAZIONE RATEI MATURATI", "205": "QUOTA A CONIUGE DIVORZIATO", "26": "INTEGRAZIONI SALARIALI", "30": "INDENNITA' DI DISOCCUPAZIONE", "35": "PRESTAZIONI VARIE", "198": "DISOCC. AGRICOLA", "196": "FONDO GARANZIA-TFR", "249": "FONDO VOLO - PREST. SOST. R."}

# ============ dagli articoli alla materia ============
# Gli avvocati non scrivono la tassonomia di SISCO: citano gli articoli.
# L'articolo e' molto piu' affidabile della prosa, che e' volutamente
# contorta. Attenzione a distinguere il RITO dalla MATERIA: l'art. 442 cpc
# dice solo che siamo in previdenza, non di che si tratta.
def _legge(num):
    """Le leggi si citano in molti modi: 'L. 222/84', 'legge n. 222 del 1984',
    'legge 222/1984'. Una sola formula per tutti."""
    return (r"l(?:egge|\.)?\s*(?:n\.?\s*)?" + str(num) +
            r"\s*(?:[\/,]|\s+del\s+|\s+dell['\u2019]anno\s+)?\s*(?:19|20)?\d{2}")

ARTICOLI_MATERIA = [
    # --- accertamento tecnico e sanitario ---
    (r"445\s*bis",                          "INVALIDITÀ CIVILE", "ATP"),
    (r"696\s*bis",                          "INVALIDITÀ CIVILE", "ATP"),
    (r"\b696\b(?!\s*bis)",                  "INVALIDITÀ CIVILE", "ATP"),
    # --- invalidita' civile e assistenza ---
    (_legge(118), "INVALIDITÀ CIVILE", None),
    (_legge(18),  "INVALIDITÀ CIVILE", "INDENNITA' ACCOMPAGNAMENTO"),
    (_legge(508), "INVALIDITÀ CIVILE", None),
    (_legge(104), "INVALIDITÀ CIVILE", None),
    (_legge(381), "INVALIDITÀ CIVILE", "SORDOMUTI"),
    (_legge(382), "INVALIDITÀ CIVILE", "CIECHI"),
    # --- invalidita' previdenziale ---
    # l'articolo puo' stare prima o dopo la legge: "art. 2 L. 222/84" e
    # "legge 222/1984 art. 2" sono la stessa citazione
    (r"(?:art\.?\s*2\b.{0,30}?222.{0,14}?(?:19)?84|"
     r"222.{0,14}?(?:19)?84.{0,30}?art\.?\s*2\b|pensione\s+(?:di\s+)?inabilit)",
     "PENSIONE INVALIDITA'", "PENSIONE INABILITA'"),
    (r"(?:art\.?\s*1\b.{0,30}?222.{0,14}?(?:19)?84|"
     r"222.{0,14}?(?:19)?84.{0,30}?art\.?\s*1\b|assegno\s+ordinario\s+(?:di\s+)?invalidit)",
     "ASSEGNO INVALIDITA'/PENSIONE INABILITA'", "ASSEGNO ORDINARIO INVALIDITA'"),
    (_legge(222), "ASSEGNO INVALIDITA'/PENSIONE INABILITA'", None),
    # --- opposizioni esecutive: avvisi di addebito, cartelle, ipoteche ---
    (r"\b615\b",                            "CONTENZIOSO CONTRIBUTIVO", "OPPOSIZ. ESECUZIONE"),
    (r"\b617\b",                            "CONTENZIOSO CONTRIBUTIVO", "OPPOSIZ. ATTI ESECUTIVI"),
    (r"\b619\b",                            "CONTENZIOSO CONTRIBUTIVO", "OPPOSIZ. DI TERZO"),
    (r"d\.?\s*lgs\.?\s*46[\/ ]?(?:19)?99|art\.?\s*24[^.]{0,20}46[\/ ]?(?:19)?99",
                                            "CONTENZIOSO CONTRIBUTIVO", "OPPOSIZ. CARTELLA"),
    (r"art\.?\s*29[^.]{0,30}78[\/ ]?2010",  "CONTENZIOSO CONTRIBUTIVO", "OPPOSIZ. AVVISO ADDEBITO"),
    (_legge(689), "CONTENZIOSO CONTRIBUTIVO", "SANZIONI AMMINISTRATIVE"),
    (r"\b77\b[^.]{0,20}602|iscrizione\s+ipotecaria", "CONTENZIOSO CONTRIBUTIVO", "IPOTECA"),
    (r"\b86\b[^.]{0,20}602|fermo\s+amministrativo",  "CONTENZIOSO CONTRIBUTIVO", "FERMO AMMINISTRATIVO"),
    # --- indebito ---
    (r"\b2033\b",                           "INDEBITO", "RIPETIZIONE INDEBITO"),
    (r"\b2946\b|\b2948\b",                  "INDEBITO", "PRESCRIZIONE"),
    (r"art\.?\s*52[^.]{0,30}88[\/ ]?(?:19)?89", "INDEBITO", None),
    # --- ammortizzatori e sussidi ---
    (r"d\.?\s*lgs\.?\s*22[\/ ]?2015|naspi",  "PRESTAZIONI - AMMORTIZZATORI SOCIALI", "NASPI"),
    (r"\bdis[\- ]?coll\b",                   "PRESTAZIONI - AMMORTIZZATORI SOCIALI", "DIS-COLL"),
    (r"cassa\s+integrazione|\bcig[so]?\b",   "PRESTAZIONI - AMMORTIZZATORI SOCIALI", "CIG"),
    (r"d\.?\s*lgs\.?\s*151[\/ ]?2001|maternit",
                                             "PRESTAZIONI - AMMORTIZZATORI SOCIALI", "MATERNITA'"),
    # --- inclusione sociale ---
    (r"d\.?\s*l(?:gs)?\.?\s*4[\/ ]?2019|reddito\s+di\s+cittadinanza|\brdc\b",
                                             "INCLUSIONE SOCIALE", "R.D.C."),
    (r"d\.?\s*l\.?\s*48[\/ ]?2023|assegno\s+di\s+inclusione|\badi\b",
                                             "INCLUSIONE SOCIALE", "ASSEGNO DI INCLUSIONE"),
    (r"assegno\s+sociale|\b" + _legge(335),
                                             "INCLUSIONE SOCIALE", "ASSEGNO SOCIALE"),
    # --- TFR e fondo di garanzia ---
    (r"\b2120\b|\bt\.?f\.?r\.?\b",           "TFS TFR", "TFR"),
    (_legge(297) + r"|fondo\s+di\s+garanzia", "TFS TFR", "FONDO DI GARANZIA"),
    # --- pensioni ---
    (r"\b1338[\/ ]?(?:19)?62|rendita\s+vitalizia", "PRESTAZIONI PENSIONISTICHE", "RENDITA VITALIZIA"),
    (r"ricongiunzion|\b" + _legge(29), "PRESTAZIONI PENSIONISTICHE", "RICONGIUNZIONE"),
    (r"riscatt",                             "PRESTAZIONI PENSIONISTICHE", "RISCATTO"),
    (r"supplement",                          "PRESTAZIONI PENSIONISTICHE", "SUPPLEMENTO"),
    (r"ricostituzion",                       "PRESTAZIONI PENSIONISTICHE", "RICOSTITUZIONE"),
    (r"pensione\s+(?:di\s+)?vecchiaia",      "PRESTAZIONI PENSIONISTICHE", "PENSIONE VO"),
    (r"pensione\s+(?:ai\s+)?superstit|reversibil", "PRESTAZIONI PENSIONISTICHE", "SUPERSTITI"),
    # --- agricoltura ---
    (r"elenchi\s+(?:nominativi\s+)?(?:dei\s+)?lavoratori\s+agricol|braccianti",
                                             "PREVIDENZA AGRICOLA", "ELENCHI"),
    # --- surroga e danni ---
    (r"\b1916\b|surrog",                     "AZIONE SURROGATORIA", None),
    (r"\b2043\b|\b2059\b|risarcimento\s+dann", "RISARCIMENTO DANNI", None),
    # --- concorsuali ---
    (r"insinuazion|stato\s+passivo|fallimen|liquidazione\s+giudizial|concordat",
                                             "PROCEDURE CONCORSUALI", None),
    # --- privacy e accesso ---
    (r"\b241[\/ ]?(?:19)?90|accesso\s+agli\s+atti", "PRIVACY/ACCESSO", "ACCESSO AGLI ATTI"),
    (r"gdpr|679[\/ ]?2016|dati\s+personali",  "PRIVACY/ACCESSO", "PRIVACY"),
]

# citano il rito, non la materia: da soli non bastano
ARTICOLI_RITO = re.compile(
    r"\b(?:442|444|409|414|702|700|281|633|640|186)\b|"
    r"art\.?\s*38\s+cost", re.I)

def materia_dagli_articoli(t):
    """Non basta trovare un articolo: in un fascicolo ne compaiono molti, e
    quello citato di sfuggita non e' la materia del giudizio. Vince quello
    che compare piu' volte, piu' presto, e vicino alle parole dell'atto
    introduttivo."""
    tt = testo(t)
    trovati = []
    for pat, materia, argomento in ARTICOLI_MATERIA:
        occ = list(re.finditer(pat, tt, re.I))
        if not occ:
            continue
        peso = len(occ) * 2
        if occ[0].start() < 3000:
            peso += 3          # nell'intestazione o nell'oggetto
        contorno = tt[max(0, occ[0].start() - 120): occ[0].end() + 120]
        if re.search(r"ricorso|ricorre|oggetto|chiede|domanda|impugna|opposizione",
                     contorno, re.I):
            peso += 4          # e' l'atto introduttivo, non una citazione di passaggio
        trovati.append((peso, len(occ), materia, argomento,
                        re.sub(r"\s+", " ", occ[0].group(0)).strip()[:40]))
    if trovati:
        trovati.sort(key=lambda x: -x[0])
        peso, n, materia, argomento, art = trovati[0]
        # incerta solo se un altro articolo ha un peso comparabile: avere
        # concorrenti deboli non e' un dubbio
        # due voci della stessa materia non sono un dubbio: lo sono solo
        # articoli che portano a materie DIVERSE
        vicini = sum(1 for x in trovati[1:] if x[0] >= peso - 2 and x[2] != materia)
        return {"materia": materia, "argomento": argomento, "articolo": art,
                "peso": peso, "occorrenze": n, "concorrenti": vicini,
                "secondo": next((x[2] for x in trovati[1:]
                                 if x[0] >= peso - 2 and x[2] != materia), None)}
    if ARTICOLI_RITO.search(tt):
        return {"materia": None, "argomento": None, "articolo": "solo il rito"}
    return None

# --------------------------------------------------------------------------
# COSA C'E' DENTRO OGNI DOCUMENTO DELLA CARTELLA
# Il nome del file non basta: spesso il decreto di fissazione udienza e'
# unito al ricorso in un PDF solo, e dalle prime quattro pagine non si vede.
# Qui ogni PDF si legge per intero e si cercano i marcatori di ogni atto; il
# risultato va in atti_letti.json nella cartella, e il registro lo legge.
# I file gia' letti e non cambiati (stessa misura e data) non si rileggono.
# --------------------------------------------------------------------------
MARCATORI_ATTI = {
    "dfu": re.compile(
        r"decreto\s+di\s+fissazione|fissa\s+(?:per\s+la\s+discussione\s+)?l['\u2019\s]*\s*udienza|"
        r"fissa\s+per\s+la\s+(?:discussione|comparizione)|visto\s+l['\u2019]?\s*art\.?\s*415|"
        r"udienza\s+di\s+discussione\s+(?:del|il|per\s+il)\s+\d", re.I),
    "relata": re.compile(r"relata\s+di\s+notific|relazione\s+di\s+notific|attestazione\s+di\s+conformit", re.I),
    "procura": re.compile(r"procura\s+(?:alle\s+liti|speciale)|conferisc\w*\s+(?:mandato|procura)|"
                          r"deleg\w*\s+a\s+rappresentar", re.I),
    "ricorso": re.compile(r"ricorso\s+(?:ex|ai\s+sensi)|giudice\s+del\s+lavoro[^\n]{0,40}\s+ricorre|"
                          r"\bricorre\b|voglia\s+l['\u2019]?\s*ill|per\s+questi\s+motivi", re.I),
}
MAX_PAGINE_ATTI = 80

def atti_nel_testo(t):
    """Gli atti riconosciuti nel testo e, per il decreto, l'udienza fissata:
    la data si cerca subito dopo il marcatore, non in tutto il documento."""
    atti, udienza = [], None
    for k, rx in MARCATORI_ATTI.items():
        m = rx.search(t or "")
        if not m:
            continue
        atti.append(k)
        if k == "dfu":
            for m2 in rx.finditer(t):
                u = trova_udienza(t[m2.start(): m2.end() + 700])
                if u:
                    u = u.isoformat() if hasattr(u, "isoformat") else str(u)
                    udienza = max(udienza or u, u)
    return atti, udienza

def leggi_atti_cartella(cart):
    try:
        from pypdf import PdfReader
    except ImportError:
        try:
            from PyPDF2 import PdfReader
        except ImportError:
            return {}
    dove = cart / "atti_letti.json"
    prima = (leggi_json(dove, {}) or {}).get("file") or {}
    out, cambiato = {}, False
    for x in sorted(cart.rglob("*")):
        if not x.is_file() or not re.search(r"\.pdf(\.p7m)?$", x.name, re.I):
            continue
        st = x.stat()
        firma = f"{st.st_size}|{int(st.st_mtime)}"
        v = prima.get(x.name)
        if v and v.get("firma") == firma:
            out[x.name] = v
            continue
        try:
            sorgente = str(x)
            if x.name.lower().endswith(".p7m"):
                dentro = apri_p7m(x.read_bytes())
                if not dentro:
                    continue
                import tempfile as _tf
                tmp = Path(_tf.gettempdir()) / ("_atti_%d.pdf" % abs(hash(str(x))))
                tmp.write_bytes(dentro)
                sorgente = str(tmp)
            r = PdfReader(sorgente)
            n = len(r.pages)
            t = "\n".join((r.pages[i].extract_text() or "") for i in range(min(n, MAX_PAGINE_ATTI)))
        except Exception:
            continue
        atti, ud = atti_nel_testo(t)
        out[x.name] = {"firma": firma, "pagine": n, "atti": atti, "udienza": ud,
                       "senza_testo": not t.strip()}
        cambiato = True
    if cambiato or set(out) != set(prima):
        try:
            scrivi_json(dove, {"letto": datetime.now().isoformat(timespec="seconds"),
                               "versione": VERSIONE, "file": out})
        except Exception as e:
            err(f"atti_letti.json in {cart.name}: {e}")
    return out

def leggi_cartella_fascicolo(cart):
    """Le stesse regole usate per le PEC, applicate ai documenti gia' in
    cartella: XML della busta prima, poi il testo dei PDF. Qui l'origine e'
    gia' nota, quindi resta solo da leggere."""
    esito = {"file": [], "note": [], "doppioni": 0, "testo": ""}
    dati = {}

    # 1) gli XML: il processo telematico li lascia dentro la cartella
    for x in sorted(cart.rglob("*.xml")):
        if x.name.lower() == "daticert.xml":
            d = leggi_daticert(x.read_bytes())
            if d.get("mittente"):
                dati.setdefault("mittente_pec", d["mittente"])
            continue
        try:
            d = leggi_xml_atto(x.read_bytes())
            for k, v in (d or {}).items():
                dati.setdefault(k, v)
        except Exception:
            pass

    # 2) il testo delle PEC conservate come .eml o .msg
    for x in sorted(list(cart.rglob("*.eml")) + list(cart.rglob("*.msg"))):
        try:
            if x.suffix.lower() == ".msg":
                dentro, corpo, err_m = estrai_da_msg(str(x))
                if corpo:
                    esito["testo"] += "\n" + corpo
            else:
                esplodi_eml(x.read_bytes(), None, set(), esito, prof=9)
        except Exception:
            pass

    # 3) il modulo di cancelleria, se c'e'
    trib = leggi_comunicazione_tribunale(esito["testo"])
    for k, v in (trib or {}).items():
        if v not in (None, "", False):
            dati.setdefault(k, v)

    # 4) i PDF: ricorso e decreto per primi, che e' dove stanno i dati
    testo_doc, nota = testo_dei_documenti(cart, max_pagine=4, max_file=12, ricorsivo=True)
    if nota:
        esito["note"].append(nota)
    tutto = (testo_doc or "") + "\n" + (esito["testo"] or "")

    rif = riferimenti(tutto[:25000])
    for k, v in rif.items():
        dati.setdefault(k, v)
    if not dati.get("parte"):
        r = trova_ricorrente(testo_doc or "", "")
        if not r:
            # come per le PEC: se nelle prime pagine il nome non c'e', il documento intero
            pieno, _n = testo_dei_documenti(cart, max_pagine=MAX_PAGINE_ATTI, max_file=12, ricorsivo=True)
            if pieno and len(pieno) > len(testo_doc or ""):
                r = trova_ricorrente(pieno, "")
        if r:
            dati["parte"] = r
    # ogni PDF letto per intero: il decreto unito al ricorso si trova qui
    atti = leggi_atti_cartella(cart)
    dati["_atti"] = atti
    ud_dfu = max((v.get("udienza") for v in atti.values() if v.get("udienza")), default=None)
    if ud_dfu:
        dati["udienza"] = ud_dfu          # il decreto fa fede: e' la data fissata
    if not dati.get("udienza"):
        u = trova_udienza(tutto)
        if u:
            dati["udienza"] = u.isoformat()
    # giudice e oggetto: nel modulo stanno in un campo, nei ricorsi nel testo
    if not dati.get("giudice"):
        # nei provvedimenti la firma, nelle relate il corpo: si prova prima
        # il corpo perche' in cartella i documenti sono quasi tutti relate
        del SCARTATI[:]
        g = giudice_dal_testo(tutto) or giudice_dalla_firma(tutto)
        dati["_scartati"] = list(SCARTATI)
        SCARTATI_TUTTI.extend(SCARTATI)
        if g:
            dati["giudice"] = g
            dati["_giudice_noto"] = any(x in GIUDICI_NOTI for x in g.split())
    a = avvocato_dal_testo(tutto, dati.get("_scartati"))
    if a:
        dati["avvocato_controparte"] = a
    if not dati.get("oggetto"):
        o = oggetto_dal_testo(tutto)
        if o:
            dati["oggetto"] = o
    # l'articolo citato vale piu' della prosa: e' li' che si legge la materia
    art = materia_dagli_articoli(tutto)
    if art:
        dati["articolo"] = art["articolo"]
        dati["_certezza"] = art.get("peso")
        dati["_concorrenti"] = art.get("concorrenti")
        dati["_secondo"] = art.get("secondo")
        if art["materia"]:
            dati["materia_sisco"] = art["materia"]
            dati["argomento_sisco"] = art["argomento"]
    # la sede va riconosciuta, non copiata: nell'XML e' "TRIBUNALE ORDINARIO
    # DI SIRACUSA", a noi serve "SIRACUSA"
    uff = sede_e_grado([tutto, dati.get("sede") or "", cart.name],
                       dati.get("mittente_pec") or "")
    if uff.get("sede"):
        dati["sede"] = uff["sede"]
        dati["grado"] = uff.get("grado") or dati.get("grado")
    elif dati.get("sede"):
        d2 = sede_e_grado([dati["sede"]], "")
        if d2.get("sede"):
            dati["sede"] = d2["sede"]
            dati["grado"] = d2.get("grado") or dati.get("grado")

    tuttif = [x for x in cart.rglob("*") if x.is_file()]
    dati["_documenti"] = len(tuttif)
    dati["_pdf"] = sum(1 for x in tuttif if x.name.lower().endswith(".pdf"))
    dati["_p7m"] = sum(1 for x in tuttif if x.name.lower().endswith(".p7m"))
    dati["_testo_letto"] = bool((testo_doc or "").strip())
    # ne conservo inizio e fine: la firma sta in fondo
    dati["_testo"] = (tutto[:3000] + "\n[...]\n" + tutto[-3000:]) if len(tutto) > 6000 else tutto
    dati["_testo_intero"] = tutto
    return dati

ESEMPI_TESTO = []
MISURE = []
SCARTATI = []
SCARTATI_TUTTI = []

def leggi_cartelle(cfg):
    global ESEMPI_TESTO, MISURE, SCARTATI, SCARTATI_TUTTI
    ESEMPI_TESTO = []
    MISURE = []
    SCARTATI = []
    SCARTATI_TUTTI = []
    """Scorre l'archivio dei fascicoli, legge i documenti e propone i dati
    mancanti in un foglio che il registro sa reimportare. Non scrive nulla
    nell'archivio: propone e basta."""
    radice = cfg.get("archivio_fascicoli")
    if not radice:
        err("archivio_fascicoli non configurato in percorsi.json")
        return 1
    p = Path(radice)
    if not p.is_absolute():
        p = BASE / p
    if not p.exists():
        err(f"cartella non trovata: {p}")
        return 1

    base = leggi_json(Path(cfg.get("base_json") or ""), None) if cfg.get("base_json") else None
    if not base:
        err("base_json non configurato o illeggibile: serve per sapere cosa manca")
        return 1

    ud_per_f = {}
    for u in base.get("udienze", []):
        if u.get("data"):
            ud_per_f.setdefault(u["fascicolo_id"], []).append(u["data"])
    giudici = {g["id"]: g.get("nome") for g in base.get("anagrafiche", {}).get("giudici", [])}
    global GIUDICI_NOTI
    GIUDICI_NOTI = {p for n in giudici.values() if n
                    for p in giudice_canonico(n).split() if len(p) > 2}

    # dal nome cartella al fascicolo, come fa il registro
    per_chiave = {}
    for f in base.get("fascicoli", []):
        for d in ud_per_f.get(f["id"], []):
            try:
                a, m, g = d.split("-")
            except Exception:
                continue
            per_chiave.setdefault(_compatta(f.get("ricorrente")) + "|" + f"{g}{m}{a[2:]}", f)

    # Un fascicolo senza udienze non entra nella chiave nome+data, ma la sua
    # cartella esiste lo stesso: si aggancia per solo nominativo, purche' non
    # ce ne siano due con lo stesso nome.
    per_nome = {}
    for f in base.get("fascicoli", []):
        n = _compatta(f.get("ricorrente"))
        if len(n) >= 5:
            per_nome.setdefault(n, []).append(f)
    solo_nome = {n: l[0] for n, l in per_nome.items() if len(l) == 1}

    cartelle = [d for d in p.rglob("*") if d.is_dir() and _legge_nome_cartella(d.name)]
    log(f"versione del programma: {VERSIONE}")
    log(f"{len(cartelle)} cartelle di fascicolo da esaminare")

    righe, senza_fasc, letti = [], 0, 0
    per_nome_n, per_solo_nome, ambigue, non_agganciate = 0, 0, [], []
    for i, cart in enumerate(cartelle, 1):
        if i % 25 == 0:
            log(f"  {i}/{len(cartelle)}...")
        chiave = _legge_nome_cartella(cart.name)
        f = per_chiave.get(chiave)
        if not f and chiave:
            # la cartella scrive il nome per esteso, il registro in forma breve:
            # "COMMENDATORE SEBASTIANO & C. SAS" e "COMMENDATORE SAS". A parita'
            # di udienza il prefisso basta a riconoscerli.
            q, ud = chiave.split("|", 1)
            vicini = []
            for k2, f2 in per_chiave.items():
                n2, ud2 = k2.split("|", 1)
                if ud2 != ud or not n2 or not q:
                    continue
                p = min(len(n2), len(q), 8)
                if p >= 4 and n2[:p] == q[:p]:
                    vicini.append(f2)
            if len(vicini) == 1:
                f = vicini[0]
                per_nome_n += 1
            elif len(vicini) > 1:
                ambigue.append((cart.name, [x.get("ricorrente") for x in vicini]))
            else:
                # ultimo tentativo: il solo nominativo, se e' univoco. Serve per
                # i fascicoli senza udienze, che la chiave nome+data non copre.
                f2 = solo_nome.get(q)
                if not f2:
                    for n2, ff in solo_nome.items():
                        p = min(len(n2), len(q), 8)
                        if p >= 5 and n2[:p] == q[:p]:
                            f2 = ff
                            break
                if f2:
                    f = f2
                    per_solo_nome += 1
        if not f:
            senza_fasc += 1
            non_agganciate.append(cart.name)
            continue
        # Non basta guardare cosa manca: il dato letto dal documento vale anche
        # quando il fascicolo ne ha gia' uno DIVERSO. Quello viene dall'Excel,
        # questo dall'atto: e' il secondo a fare fede.
        manca = {
            "giudice": not f.get("giudice_id"),
            "materia": not f.get("materia_testo"),
            "rg": not f.get("rg"),
            "sede": not f.get("sede"),
            "udienza": not ud_per_f.get(f["id"]),
        }
        d = leggi_cartella_fascicolo(cart)
        letti += 1
        tx = d.get("_testo_intero") or ""
        MISURE.append({
            "car": len(tx),
            "testo": bool(tx.strip()),
            "giudice_parola": bool(re.search(r"giudice", tx, re.I)),
            "dott": bool(re.search(r"\bdott|\bdr\.", tx, re.I)),
            "oggetto_parola": bool(re.search(r"oggetto", tx, re.I)),
            "art": bool(re.search(r"art\.?\s*\d|l(?:egge)?\.?\s*n?\.?\s*\d+[\/ ]\d", tx, re.I)),
            "trovato_giudice": bool(d.get("giudice")),
            "trovato_materia": bool(d.get("materia_sisco") or d.get("oggetto")),
        })
        if len(ESEMPI_TESTO) < 12 and not d.get("giudice") and \
           len((d.get("_testo") or "").strip()) > 400:
            ESEMPI_TESTO.append((cart.name, d["_testo"]))
        prop, diverso = {}, []
        gia_g = giudice_canonico(giudici.get(f.get("giudice_id")) or "")
        letto_g = giudice_canonico(d.get("giudice") or "")
        if letto_g and (manca["giudice"] or
                        (gia_g and letto_g.split()[0] != gia_g.split()[0])):
            prop["GIUDICE"] = letto_g
            if not manca["giudice"]:
                diverso.append(f"giudice: nel registro {gia_g}")
        if d.get("materia_sisco"):
            # la materia dedotta dall'articolo e' della tassonomia SISCO:
            # vale anche in sostituzione di quella scritta a mano
            gia_m = testo(f.get("materia_testo"))
            if manca["materia"] or _compatta(gia_m) != _compatta(d["materia_sisco"]):
                prop["MATERIA"] = d["materia_sisco"]
                if d.get("argomento_sisco"):
                    prop["ARGOMENTO"] = d["argomento_sisco"]
                if not manca["materia"]:
                    diverso.append(f"materia: nel registro {gia_m}")
        elif manca["materia"] and (d.get("provvedimento") or d.get("oggetto")):
            prop["MATERIA"] = (d.get("oggetto") or d.get("provvedimento"))[:60]
        if d.get("avvocato_controparte") and not f.get("avvocato_controparte_id"):
            prop["AVV. CONTROPARTE"] = d["avvocato_controparte"]
        if d.get("rg") and (manca["rg"] or
                            _compatta(f.get("rg")) != _compatta(d["rg"])):
            prop["RG"] = d["rg"]
            if not manca["rg"]:
                diverso.append(f"ruolo: nel registro {f.get('rg')}")
        if manca["sede"] and d.get("sede"):
            prop["SEDE"] = d["sede"]
            prop["GRADO"] = d.get("grado") or ""
        if manca["udienza"]:
            # la data e' nel nome della cartella: "BUFALINO C UD 030726".
            # E' piu' affidabile di quella letta dai documenti, perche' e'
            # l'ufficio ad averla scritta.
            dal_nome = _legge_nome_cartella(cart.name)
            u = None
            if dal_nome and "|" in dal_nome:
                gg = dal_nome.split("|")[1]
                if len(gg) == 6:
                    u = f"{gg[0:2]}/{gg[2:4]}/20{gg[4:6]}"
            if not u and d.get("udienza"):
                try:
                    a, m, g = str(d["udienza"])[:10].split("-")
                    u = f"{g}/{m}/{a}"
                except Exception:
                    pass
            if u:
                prop["DATA UDIENZA"] = u
        if not prop:
            continue
        righe.append({
            "CODICE": f["id"],
            "GIUDICE": prop.get("GIUDICE", giudici.get(f.get("giudice_id")) or ""),
            "MATERIA": prop.get("MATERIA", f.get("materia_testo") or ""),
            "FASCICOLO": f.get("ricorrente") or "",
            "DATA UDIENZA": prop.get("DATA UDIENZA", ""),
            "DATA PEC": "",
            "DIFENSORE INPS": "",
            "DATA COSTITUZIONE": "",
            "SISCO": f.get("sisco") or "",
            "RG": prop.get("RG", f.get("rg") or ""),
            "SEDE": prop.get("SEDE", f.get("sede") or ""),
            "GRADO": prop.get("GRADO", f.get("grado") or ""),
            "AVV. CONTROPARTE": prop.get("AVV. CONTROPARTE", ""),
            "STATO": f.get("stato") or "",
            "ARGOMENTO": prop.get("ARGOMENTO", ""),
            "NOTE": ("SOLO IL SISCO DA CARICARE - " if not f.get("sisco") and
                     prop.get("GIUDICE") and prop.get("MATERIA") and
                     (prop.get("RG") or f.get("rg")) else "") +
                    "letto dai documenti in cartella: " + ", ".join(prop.keys()) +
                    (" [art. " + d["articolo"] +
                     (", ma pesa quanto %s: DA CONTROLLARE" % d["_secondo"]
                      if d.get("_concorrenti") else "") + "]"
                     if d.get("articolo") and prop.get("MATERIA") else "") +
                    ((" - GIUDICE DA CONFERMARE: non e' fra quelli della sezione, "
                      "potrebbe essere un avvocato") if prop.get("GIUDICE") and
                     not d.get("_giudice_noto") else "") +
                    ((" - DIVERSO DA QUELLO IN ARCHIVIO (" + "; ".join(diverso) + ")")
                     if diverso else "") +
                    (" (nessun testo estratto dai PDF: forse scansioni)"
                     if not d.get("_testo_letto") else ""),
            "_cartella": cart.name,
            "_documenti": d.get("_documenti", 0),
            "_secondo": d.get("_secondo"),
            "_concorrenti": d.get("_concorrenti"),
            "_pdf": d.get("_pdf", 0),
            "_p7m": d.get("_p7m", 0),
        })

    if not righe:
        log("")
        log(f"{letti} cartelle lette: nessun dato nuovo da proporre.")
        return 0

    # se un dato non si trova quasi mai, il problema non e' nei documenti ma
    # nelle formule che cerco: mettere sotto gli occhi il testo vero e' l'unico
    # modo per correggerle
    campioni = []
    for cart, testo_visto in ESEMPI_TESTO[:12]:
        estratto = []
        # la firma sta in fondo: e' li' che va guardato, non solo dove
        # compare la parola "giudice" nel corpo del ricorso
        coda = testo_visto[-1200:]
        if coda.strip():
            estratto.append("[ULTIME RIGHE] " + re.sub(r"\s+", " ", coda[-420:]).strip())
        for parola in ("il giudice", "f.to", "firmato", "dott", "oggetto"):
            for m in re.finditer(re.escape(parola), testo_visto, re.I):
                a = max(0, m.start() - 30)
                fr = re.sub(r"\s+", " ", testo_visto[a:m.start() + 90]).strip()
                if fr not in estratto:
                    estratto.append(fr)
                break
            if len(estratto) >= 5:
                break
        if estratto:
            campioni.append((cart, estratto))

    # accanto al programma, non in IN ARRIVO: quella cartella viene svuotata
    # da --azzera, e questo file serve dopo
    dest = BASE
    dest.mkdir(parents=True, exist_ok=True)
    col = ["CODICE", "GIUDICE", "MATERIA", "ARGOMENTO", "FASCICOLO", "DATA UDIENZA",
           "DATA PEC", "DIFENSORE INPS", "DATA COSTITUZIONE", "SISCO", "RG", "SEDE",
           "GRADO", "AVV. CONTROPARTE", "STATO", "NOTE"]
    q = lambda v: '"' + str(v if v is not None else "").replace('"', '""') + '"'
    C = ["\ufeff" + ";".join(q(c) for c in col)]
    for r in righe:
        C.append(";".join(q(r.get(c, "")) for c in col))
    (dest / "_DATI_DAI_DOCUMENTI.csv").write_text("\r\n".join(C), encoding="utf-8")

    # Misuro il testo prima di cercarci dentro: se le parole chiave non ci sono,
    # nessuna formula potra' trovarle, e continuare a correggerle e' inutile.
    if MISURE:
        tot = len(MISURE)
        m = {k: sum(1 for x in MISURE if x.get(k)) for k in
             ("testo", "art", "giudice_parola", "oggetto_parola", "dott",
              "trovato_giudice", "trovato_materia")}
        car = sorted(x["car"] for x in MISURE)
        log("")
        log("MISURE SUL TESTO LETTO (su %d cartelle):" % tot)
        log("  %5d  hanno testo leggibile" % m["testo"])
        log("  %5d  contengono la parola 'giudice'" % m["giudice_parola"])
        log("  %5d  contengono 'dott' o 'dr.'" % m["dott"])
        log("  %5d  contengono la parola 'oggetto'" % m["oggetto_parola"])
        log("  %5d  contengono un articolo di legge riconoscibile" % m["art"])
        log("  %5d  -> giudice effettivamente estratto" % m["trovato_giudice"])
        log("  %5d  -> materia effettivamente estratta" % m["trovato_materia"])
        if car:
            log("  caratteri letti per cartella: minimo %d, mediana %d, massimo %d"
                % (car[0], car[len(car)//2], car[-1]))

    conta = {}
    for r in righe:
        for k in ("GIUDICE", "MATERIA", "ARGOMENTO", "RG", "SEDE", "DATA UDIENZA"):
            if k in r["NOTE"]:
                conta[k] = conta.get(k, 0) + 1
    correzioni = sum(1 for r in righe if "DIVERSO DA QUELLO" in r["NOTE"])
    pronti = sum(1 for r in righe if "SOLO IL SISCO" in r["NOTE"])
    avv_nuovi = sum(1 for r in righe if r.get("AVV. CONTROPARTE"))
    gius = {}
    for r in righe:
        if r.get("GIUDICE") and "GIUDICE" in r["NOTE"]:
            gius[r["GIUDICE"]] = gius.get(r["GIUDICE"], 0) + 1
    incerte = sum(1 for r in righe if "DA CONTROLLARE" in r["NOTE"])
    muti = sum(1 for r in righe if "scansioni" in r["NOTE"])
    tot_pdf = sum(r.get("_pdf", 0) for r in righe)
    tot_p7m = sum(r.get("_p7m", 0) for r in righe)

    L = ["DATI RICAVATI DAI DOCUMENTI - " + datetime.now().strftime("%d/%m/%Y %H:%M")
         + "   [programma: " + VERSIONE + "]",
         "=" * 100,
         f"{letti} cartelle lette, {len(righe)} fascicoli con dati nuovi da proporre.",
         f"{senza_fasc} cartelle senza un fascicolo corrispondente (non toccate)."] + \
        ([f"{per_nome_n} riconosciute dal nome abbreviato, a parita' di udienza."]
         if per_nome_n else []) + \
        ([f"{per_solo_nome} riconosciute dal solo nominativo: sono fascicoli senza udienze,"
          " che hanno comunque la loro cartella."] if per_solo_nome else []) + \
        ([f"{len(ambigue)} cartelle che corrispondono a piu' fascicoli: lasciate stare."]
         if ambigue else []) + \
        ["",
         "",
         "COSA E' STATO TROVATO:"] + \
        [f"  {n:5d}  {k.lower()}" for k, n in sorted(conta.items(), key=lambda x: -x[1])] + \
        ([f"  {muti:5d}  cartelle senza testo leggibile nei PDF: sono scansioni, "
          "servirebbe l'OCR"] if muti else []) + \
        ([f"  {correzioni:5d}  di questi correggono un dato gia' in archivio, "
          "che risulta diverso"] if correzioni else []) + \
        ([f"  {incerte:5d}  materie incerte: piu' articoli concorrono, "
          "sono marcate DA CONTROLLARE"] if incerte else []) + \
        ([f"  {pronti:5d}  hanno tutto tranne il SISCO: pronti da caricare in piattaforma",
          "         (filtra 'SOLO IL SISCO' nella colonna NOTE)"] if pronti else []) + \
        ([f"  {avv_nuovi:5d}  avvocati di controparte, letti dalla firma della relata"]
         if avv_nuovi else []) + \
        ([""] + ["GIUDICI LETTI DAI DOCUMENTI:"] +
         [f"  {n:5d}  {g}" for g, n in sorted(gius.items(), key=lambda x: -x[1])] +
         ["  Solo i giudici in anagrafica: gli altri nomi sono avvocati e",
          "  sono finiti nella colonna AVV. CONTROPARTE."]
         if gius else []) + \
        ([""] + ["NOMI SCARTATI COME GIUDICI (non sono in anagrafica):"] +
         [f"  {n:5d}  {g}" for g, n in sorted(
             {x: SCARTATI_TUTTI.count(x) for x in set(SCARTATI_TUTTI)}.items(),
             key=lambda x: -x[1])[:15]] +
         ["  Sono trattati come avvocati. Se fra questi c'e' un magistrato",
          "  in servizio, aggiungilo dal registro: Cartelle e manutenzione."]
         if SCARTATI_TUTTI else []) + \
        ["", f"  documenti esaminati: {tot_pdf} PDF e {tot_p7m} firmati .p7m"] + \
        ([""] + ["MISURE SUL TESTO LETTO (perche' un dato non si trova):",
          f"  {sum(1 for x in MISURE if x['testo']):5d} su {len(MISURE)} hanno testo leggibile",
          f"  {sum(1 for x in MISURE if x['giudice_parola']):5d} contengono la parola 'giudice'",
          f"  {sum(1 for x in MISURE if x['dott']):5d} contengono 'dott' o 'dr.'",
          f"  {sum(1 for x in MISURE if x['oggetto_parola']):5d} contengono la parola 'oggetto'",
          f"  {sum(1 for x in MISURE if x['art']):5d} contengono un articolo di legge",
          f"  {sum(1 for x in MISURE if x['trovato_giudice']):5d} -> giudice estratto",
          f"  {sum(1 for x in MISURE if x['trovato_materia']):5d} -> materia estratta",
          "  Se la parola c'e' ma il dato non si estrae, il difetto e' nella",
          "  formula. Se la parola non c'e', il dato in quei documenti non c'e'."]
         if MISURE else []) + \
        [
         "",
         "Il file _DATI_DAI_DOCUMENTI.csv si importa direttamente dal registro:",
         "Cartelle e manutenzione \u2192 Rileggi il foglio. Se vuoi controllarlo",
         "prima, aprilo con Excel e correggi quello che serve.",
         "",
         "Qui NON e' stato scritto nulla nell'archivio: sono proposte."] + \
        ([""] + ["=" * 100,
         "COM'E' SCRITTO NEI TUOI DOCUMENTI (dove non ho trovato il giudice)",
         "=" * 100,
         "Se il giudice c'e' ma non lo riconosco, qui sotto si vede come e'",
         "scritto: mandalo a chi cura il programma e la formula viene aggiunta.",
         ""] + [r for c, ex in campioni
                for r in [f"  {c[:70]}"] + [f"      ...{e}..." for e in ex] + [""]]
         if campioni else []) + \
        ["",
         f"{'FASCICOLO':<30}{'COSA HO TROVATO':<52}DOC",
         "-" * 100]
    for r in righe[:200]:
        trovato = [f"{k}={r[k]}" for k in ("GIUDICE", "MATERIA", "RG", "SEDE", "DATA UDIENZA")
                   if r.get(k) and k in r["NOTE"]]
        L.append(f"{testo(r['FASCICOLO'])[:29]:<30}{', '.join(trovato)[:51]:<52}{r['_documenti']:>3}")
    (dest / "_DATI_DAI_DOCUMENTI.txt").write_text("\n".join(L), encoding="utf-8")

    log("")
    log(f"LETTE {letti} cartelle, {len(righe)} fascicoli con dati da proporre.")
    log(f"  {dest / '_DATI_DAI_DOCUMENTI.csv'}")
    log("  Il registro lo importa cosi' com'e': Cartelle e manutenzione ->")
    log("  Rileggi il foglio. Se vuoi controllarlo prima, aprilo con Excel.")
    return 0

def sposta_archivio_accanto_al_registro(cfg, ricognizione=False):
    """archivio_pec.json sta accanto a base.json, nella cartella del registro:
    e' l'unica che il registro ha sempre aperta, quindi lo trova da solo.
    Nella cartella del programma andava collegato a mano, e il collegamento
    si perdeva ogni volta che il browser toglieva il permesso. Un solo file:
    quello vecchio si sposta, non si copia.

    IN RICOGNIZIONE NON SI SPOSTA NIENTE. Un giro che dichiara "nessuna
    scrittura" non puo' spostare un file: l'archivio sembrava sparito perche'
    una ricognizione lo aveva portato da un'altra parte, e chi lo cercava dove
    stava prima non lo trovava piu'. Si limita a dire dove andrebbe."""
    global ARCHIVIO
    bj = cfg.get("base_json")
    if not bj or SIMULAZIONE:
        return
    if ricognizione:
        q0 = Path(bj)
        if not q0.is_absolute():
            q0 = BASE / q0
        n0 = (q0 if q0.is_dir() else q0.parent) / "archivio_pec.json"
        if n0 != ARCHIVIO and ARCHIVIO.exists():
            log(f"(in ricognizione non lo sposto: archivio_pec.json andrebbe in {n0.parent})")
        if n0.exists():
            ARCHIVIO = n0          # leggere quello giusto si', scriverlo no
        return
    q = Path(bj)
    if not q.is_absolute():
        q = BASE / q
    nuovo = (q if q.is_dir() else q.parent) / "archivio_pec.json"
    if not nuovo.parent.exists() or nuovo == ARCHIVIO:
        return
    vecchio = ARCHIVIO
    try:
        if vecchio.exists() and not nuovo.exists():
            shutil.move(str(vecchio), str(nuovo))
            log(f"archivio_pec.json spostato accanto al registro: {nuovo}")
        elif vecchio.exists() and nuovo.exists():
            if vecchio.stat().st_mtime > nuovo.stat().st_mtime:
                shutil.move(str(vecchio), str(nuovo))      # il piu' recente vince
            else:
                vecchio.unlink()
            log(f"archivio_pec.json: tenuto uno solo, in {nuovo.parent}")
    except Exception as e:
        err(f"archivio_pec.json non spostato ({e}): resta in {vecchio}")
        return
    ARCHIVIO = nuovo

def salva_archivio(dati):
    """In simulazione la memoria dei messaggi visti non si tocca: una prova
    non deve lasciare tracce che alterino i lanci veri."""
    if SIMULAZIONE:
        return
    scrivi_json(ARCHIVIO, dati)

def converti_xls(percorso):
    """Le estrazioni di SISCO escono in .xls, il formato binario di Excel 97,
    che il registro non sa leggere: e' un contenitore OLE, non uno zip.
    Qui si converte in .xlsx, che il registro legge senza librerie."""
    p = Path(percorso)
    if not p.is_absolute():
        p = BASE / p
    if p.is_dir():
        files = sorted(p.glob("*.xls"))
        if not files:
            err(f"nessun file .xls in {p}")
            return 1
    elif p.exists():
        files = [p]
    else:
        err(f"non trovo {p}")
        return 1

    try:
        import xlrd
    except ImportError:
        log("manca la libreria xlrd: provo a installarla")
        import subprocess
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "xlrd"],
                           check=True, capture_output=True)
            import xlrd
        except Exception as e:
            err("installazione non riuscita: " + str(e)[:100])
            err("Installala a mano con:  py -m pip install xlrd")
            err("Oppure apri il file con Excel e salvalo come .xlsx.")
            return 1

    def colonna(i):
        s, n = "", i + 1
        while n > 0:
            s = chr(65 + (n - 1) % 26) + s
            n = (n - 1) // 26
        return s

    fatti = 0
    for f in files:
        try:
            w = xlrd.open_workbook(str(f))
            s = w.sheet_by_index(0)

            def cella(r, c):
                """Le date devono uscire come SERIALE, sempre.

                Il registro legge le date con parseFloat e le converte dal
                seriale di Excel. Se una cella e' testo - "02/01/2024" -
                parseFloat legge 2, che il registro scarta, e il fascicolo
                nasce senza apertura e senza udienza: e' il motivo per cui
                centinaia di fascicoli caricati da SISCO restavano muti.

                Due casi da normalizzare:
                  - cella di tipo DATE: xlrd da' gia' il seriale, si tiene;
                  - cella di tipo TEXT che CONTIENE una data: va riportata a
                    seriale, altrimenti si perde.
                """
                t = s.cell_type(r, c)
                v = s.cell_value(r, c)
                if t == xlrd.XL_CELL_DATE:
                    return v                      # gia' seriale
                if t == xlrd.XL_CELL_TEXT:
                    m = re.match(r"^\s*(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})"
                                 r"(?:\s+(\d{1,2}):(\d{2}))?\s*$", str(v))
                    if m:
                        g, me, a = int(m.group(1)), int(m.group(2)), int(m.group(3))
                        if a < 100:
                            a += 2000 if a < 70 else 1900
                        try:
                            d = date(a, me, g)
                        except ValueError:
                            return v              # non e' una data: resta testo
                        n = (d - date(1899, 12, 30)).days
                        if m.group(4):
                            n += (int(m.group(4)) * 60 + int(m.group(5))) / 1440.0
                        return float(n)
                return v

            conta = {"date": 0, "testo_convertito": 0, "testo_non_riconosciuto": {}}

            def cella_contata(r, c):
                t = s.cell_type(r, c)
                v = s.cell_value(r, c)
                out = cella(r, c)
                if t == xlrd.XL_CELL_DATE:
                    conta["date"] += 1
                elif t == xlrd.XL_CELL_TEXT and out is not v:
                    conta["testo_convertito"] += 1
                return out

            # quali celle sono date: serve per formattarle, altrimenti in
            # Excel si vede il seriale nudo (46210 invece di 07/07/2026) e il
            # file sembra rotto anche quando e' corretto
            sono_date = set()

            def cella_marcata(r, c):
                t = s.cell_type(r, c)
                v = cella_contata(r, c)
                if t == xlrd.XL_CELL_DATE or (t == xlrd.XL_CELL_TEXT
                                              and isinstance(v, float)):
                    sono_date.add((r, c))
                return v

            righe = [[cella_marcata(r, c) for c in range(s.ncols)] for r in range(s.nrows)]

            # Quali colonne SONO date, e cosa contengono davvero. Senza questo
            # si tira a indovinare: una data puo' arrivare come cella data,
            # come testo, o in un formato che non riconosciamo, e le tre cose
            # richiedono correzioni diverse.
            if righe:
                intest = [str(x).strip() for x in righe[0]]
                for j, nome in enumerate(intest):
                    if not re.search(r"data|udienza|apertura|scadenz|notific",
                                     nome, re.I):
                        continue
                    tipi, esempi = {}, []
                    for r in range(1, min(s.nrows, 200)):
                        if j >= s.ncols:
                            continue
                        t = s.cell_type(r, j)
                        v = s.cell_value(r, j)
                        tipi[t] = tipi.get(t, 0) + 1
                        if len(esempi) < 3 and v not in (None, "", 0, 1):
                            esempi.append(repr(v)[:26])
                    nomi_tipo = {0: "vuota", 1: "TESTO", 2: "numero",
                                 3: "data", 4: "bool", 5: "errore"}
                    dett = ", ".join(f"{nomi_tipo.get(k, k)}:{n}"
                                     for k, n in sorted(tipi.items()))
                    log(f"    colonna '{nome[:26]}' -> {dett}   es. {', '.join(esempi)}")
        except Exception as e:
            err(f"{f.name}: {e}")
            continue

        def x(v):
            if isinstance(v, float) and v == int(v) and abs(v) < 1e15:
                v = int(v)
            return (str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

        def cel(i, j, v):
            """Una data si scrive come numero con lo stile 1 (gg/mm/aaaa):
            cosi' Excel la mostra come data e il registro continua a leggere
            il seriale. Scritta come testo era corretta ma illeggibile."""
            rif = "%s%d" % (colonna(j), i + 1)
            if (i, j) in sono_date and isinstance(v, (int, float)) and v:
                return '<c r="%s" s="1"><v>%s</v></c>' % (rif, repr(float(v)).rstrip('0').rstrip('.'))
            return ('<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                    % (rif, x(v)))

        sd = "".join(
            '<row r="%d">%s</row>' % (i + 1, "".join(
                cel(i, j, v) for j, v in enumerate(rg)))
            for i, rg in enumerate(righe))

        parti = {
          "[Content_Types].xml":
            '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/'
            'content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats'
            '-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-'
            'officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/'
            'sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.'
            'spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'spreadsheetml.styles+xml"/></Types>',
          "_rels/.rels":
            '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package'
            '/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org'
            '/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>',
          "xl/workbook.xml":
            '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/'
            'spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
            '2006/relationships"><sheets><sheet name="Dati" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>',
          "xl/_rels/workbook.xml.rels":
            '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package'
            '/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org'
            '/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument'
            '/2006/relationships/styles" Target="styles.xml"/></Relationships>',
          "xl/styles.xml":
            '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/'
            'spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/>'
            '<name val="Calibri"/></font></fonts><fills count="1"><fill>'
            '<patternFill patternType="none"/></fill></fills><borders count="1">'
            '<border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" '
            'fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2">'
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" '
            'applyNumberFormat="1"/></cellXfs><cellStyles count="1">'
            '<cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>',
          "xl/worksheets/sheet1.xml":
            '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/'
            'spreadsheetml/2006/main"><sheetData>' + sd + '</sheetData></worksheet>',
        }
        fuori = f.with_suffix(".xlsx")   # zipfile e' gia' importato in testa
        with zipfile.ZipFile(fuori, "w", zipfile.ZIP_DEFLATED) as z:
            for n, c in parti.items():
                z.writestr(n, c)
        log(f"  {f.name}  ->  {fuori.name}   ({len(righe)} righe)")
        log(f"    celle data: {conta['date']}, "
            f"testi riconosciuti come data e convertiti: {conta['testo_convertito']}")
        fatti += 1

    if fatti:
        log("")
        log(f"{fatti} file convertiti. Ora caricali dal registro, scheda Avvio,")
        log("uno alla volta: «Carica un trimestre».")
    return 0

# ============ i fascicoli scritti negli OGGETTI ============
# I ricorsi non arrivano sempre come notifica con allegati: spesso sono
# comunicazioni interne che li elencano nell'oggetto, piu' fascicoli per volta.
# Sono la fonte piu' pulita che abbiamo, e finora la buttavamo via.

# "INVIO RICORSI SIRACUSA - ABRAMO ANGELO 533/26/CO - GIULIANO M. 535/26/CO"
RE_COPPIA = re.compile(
    r"([A-ZÀ-Ù][A-ZÀ-Ù'’\.\- ]{3,60}?)\s*[-–]?\s*"
    r"(\d{1,4}\s*/\s*\d{2}\s*/\s*(?:CO|AO|CV|AI|IC)(?:\s*/\s*\d)?)", re.I)

# "RICHIESTA ISTRUTTORIA (Sede: SIRACUSA - Pratica: 533/26/CO/1 - ABRAMO ANGELO)"
RE_MODULO = re.compile(
    r"\(\s*sede\s*:\s*([A-ZÀ-Ù ]{3,20}?)\s*[-–]\s*"
    r"pratica\s*:\s*(\d{1,4}\s*/\s*\d{2}\s*/\s*(?:CO|AO|CV|AI|IC)(?:\s*/\s*\d)?)\s*"
    r"[-–]\s*([^)]{3,60})\)", re.I)

def fascicoli_dall_oggetto(oggetto):
    """Restituisce l'elenco dei fascicoli nominati nell'oggetto.
    Un messaggio puo' portarne piu' di uno: 'INVIO RICORSI - TIZIO 533/26/CO -
    CAIO 535/26/CO' sono due ricorsi, non uno."""
    ogg = re.sub(r"\s+", " ", testo(oggetto)).strip()
    fuori, visti = [], set()

    for m in RE_MODULO.finditer(ogg):
        sede, prat, nome = m.group(1).strip(), m.group(2), m.group(3).strip()
        n = _ripulisci_nome_persona(nome)
        if n and _compatta(prat) not in visti:
            visti.add(_compatta(prat))
            fuori.append({"ricorrente": n, "sisco": re.sub(r"\s+", "", prat),
                          "sede": sede.upper(), "da": "modulo istruttoria"})
    if fuori:
        return fuori

    # senza il modulo: le coppie nome + numero di pratica
    if not re.search(r"invio\s+ricors|ricors\w*\s+(?:nuovi|assegnat)|"
                     r"trasmission\w*\s+ricors|nuovi\s+ricors", ogg, re.I) \
       and not RE_COPPIA.search(ogg):
        return []
    for m in RE_COPPIA.finditer(ogg):
        nome, prat = m.group(1).strip(" -–"), m.group(2)
        # via le parole di servizio che precedono il nome
        nome = re.sub(r"^(?:i|r|fw|fwd|re)\s*:\s*", "", nome, flags=re.I)
        # via TUTTE le parole di servizio, non solo fino alla prima: l'oggetto
        # dice "INVIO RICORSI SIRACUSA - COGNOME", e il nome e' quel che resta
        for _ in range(6):
            n2 = re.sub(r"^\s*[-–:]?\s*(?:invio|invii|trasmissione|trasm\.?|ricorso|"
                        r"ricorsi|siracusa|catania|noto|augusta|lentini|nuovi|nuove|"
                        r"assegnati|assegnate|pratica|pratiche|sede|urgente|elenco|"
                        r"allegat\w*)\b\s*[-–:.]?\s*", "", nome, flags=re.I)
            if n2 == nome:
                break
            nome = n2
        n = _ripulisci_nome_persona(nome)
        if n and len(n) >= 4 and _compatta(prat) not in visti:
            visti.add(_compatta(prat))
            fuori.append({"ricorrente": n, "sisco": re.sub(r"\s+", "", prat),
                          "sede": None, "da": "elenco nell'oggetto"})
    return fuori

def raccogli_dagli_oggetti(cfg, args=None):
    """Passa in rassegna TUTTE le PEC, anche quelle scartate, e raccoglie i
    fascicoli nominati negli oggetti. Non crea cartelle: produce un elenco da
    importare nel registro, dove i fascicoli nascono con nome e SISCO."""
    ns = collega_outlook()
    if not ns:
        return 1
    cartelle = scegli_cartelle(ns, cfg.get("cartelle", []), cfg.get("escludi_cartelle", []))
    filtro, descr = filtro_date(cfg)
    log(f"periodo: {descr}")
    trovati, esaminati = {}, 0
    for cart in cartelle:
        try:
            items = prop(cart, "Items")
            items.Sort("[ReceivedTime]", True)
            if filtro:
                items = items.Restrict(filtro)
        except Exception:
            continue
        it = items.GetFirst()
        while it is not None:
            m, it = it, items.GetNext()
            if not e_un_messaggio(m):
                continue
            esaminati += 1
            ogg = testo(prop(m, "Subject"))
            for f in fascicoli_dall_oggetto(ogg):
                k = _compatta(f["sisco"])
                if k in trovati:
                    continue
                f["quando"] = str(prop(m, "ReceivedTime", ""))[:10]
                f["oggetto"] = ogg[:110]
                f["cartella"] = testo(prop(cart, "Name"))
                trovati[k] = f
    righe = sorted(trovati.values(), key=lambda x: x["sisco"])
    log("")
    log("=" * 70)
    log(f"{len(righe)} fascicoli nominati negli oggetti, su {esaminati} messaggi")
    per_fonte = {}
    for x in righe:
        per_fonte[x["da"]] = per_fonte.get(x["da"], 0) + 1
    for k, v in sorted(per_fonte.items(), key=lambda x: -x[1]):
        log(f"  {v:5d}  {k}")
    if not righe:
        log("Nessuno: gli oggetti non contengono coppie nome + numero di pratica.")
        return 0
    col = ["CODICE", "GIUDICE", "MATERIA", "ARGOMENTO", "FASCICOLO", "DATA UDIENZA",
           "DATA PEC", "DIFENSORE INPS", "DATA COSTITUZIONE", "SISCO", "RG",
           "SEDE", "GRADO", "AVV. CONTROPARTE", "STATO", "NOTE"]
    out = BASE / "_DAGLI_OGGETTI.csv"
    with open(out, "w", newline="", encoding="utf-8") as fh:
        fh.write("\ufeff")
        w = csv.DictWriter(fh, col, delimiter=";", quoting=csv.QUOTE_ALL)
        w.writeheader()
        for x in righe:
            w.writerow({"CODICE": "", "GIUDICE": "", "MATERIA": "", "ARGOMENTO": "",
                        "FASCICOLO": x["ricorrente"], "DATA UDIENZA": "",
                        "DATA PEC": "/".join(reversed(x["quando"].split("-")))
                                    if x["quando"] else "",
                        "DIFENSORE INPS": "", "DATA COSTITUZIONE": "",
                        "SISCO": x["sisco"], "RG": "",
                        "SEDE": x["sede"] or "",
                        # 1/25/AO/0 e' un appello, .../CO/2 e' secondo grado:
                        # il numero di pratica lo dice, non serve indovinarlo
                        "GRADO": ("appello"
                                  if re.search(r"/AO/|/\s*2\s*$", x["sisco"], re.I)
                                  else "primo"),
                        "AVV. CONTROPARTE": "", "STATO": "attivo",
                        "NOTE": f"dall'oggetto di una PEC ({x['da']}): {x['oggetto']}"})
    log("")
    log(f"  {out}")
    log("  Importalo dal registro: Cartelle e manutenzione -> Rileggi il foglio.")
    log("  I fascicoli nascono con nominativo e numero SISCO; ruolo e udienza")
    log("  arrivano dopo, dai trimestri o dalle comunicazioni successive.")
    log("=" * 70)
    for x in righe[:12]:
        log(f"    {x['sisco']:<14} {x['ricorrente'][:34]:<36} {x['da']}")
    return 0

BACINO = BASE / "bacino.json"

def raccogli_bacino(cfg, args):
    """LEGGE OUTLOOK UNA VOLTA SOLA.

    Il tempo se ne va tutto qui: undicimila messaggi via COM sono dieci
    minuti. I criteri invece si applicano in un istante. Quindi si separano
    le due cose: questo comando salva in locale quello che serve a decidere —
    oggetto, mittente, destinatari, nomi degli allegati — e da quel momento
    si puo' riclassificare all'infinito senza toccare la casella."""
    ns = collega_outlook()
    if not ns:
        return 1
    cartelle = scegli_cartelle(ns, cfg.get("cartelle", []), cfg.get("escludi_cartelle", []))
    if not cartelle:
        err("nessuna cartella trovata")
        return 1
    filtro, descr = filtro_date(cfg)
    log(f"periodo: {descr}")
    log("Lettura di Outlook: e' la parte lenta, ma si fa una volta sola.")

    msg, n = [], 0
    for cart in cartelle:
        nome_c = testo(prop(cart, "Name"))
        try:
            items = prop(cart, "Items")
            items.Sort("[ReceivedTime]", True)
            if filtro:
                items = items.Restrict(filtro)
        except Exception as e:
            err(f"{nome_c}: {e}")
            continue
        quanti = 0
        it = items.GetFirst()
        while it is not None:
            m, it = it, items.GetNext()
            if not e_un_messaggio(m):
                continue
            n += 1
            quanti += 1
            if n % 1000 == 0:
                log(f"    ...{n} messaggi")
            allegati = []
            try:
                for j in range(1, m.Attachments.Count + 1):
                    allegati.append(testo(prop(m.Attachments.Item(j), "FileName")))
            except Exception:
                pass
            try:
                k = chiave_stabile(m)
            except Exception:
                k = None
            msg.append({
                "k": k,
                "cartella": nome_c,
                "percorso": testo(prop(cart, "FolderPath")),
                "data": str(prop(m, "ReceivedTime", ""))[:19],
                "oggetto": testo(prop(m, "Subject")),
                "mittente": mittente_vero(m),
                "destinatari": (str(prop(m, "To", "")) + " " + str(prop(m, "CC", "")))[:400],
                "allegati": allegati,
                "corpo": testo(prop(m, "Body"))[:1500] if cfg.get("usa_corpo") else "",
            })
        log(f"  {nome_c}  ({quanti})")

    scrivi_json(BACINO, {"aggiornato": datetime.now().isoformat(timespec="seconds"),
                         "periodo": descr, "messaggi": msg})
    log("")
    log("=" * 68)
    log(f"BACINO: {len(msg)} messaggi salvati in {BACINO.name}")
    log("Da adesso:  python pec_intake.py --riclassifica")
    log("Non tocca Outlook e risponde subito: si possono provare i criteri")
    log("quante volte serve, finche' gli scarti non convincono.")
    log("=" * 68)
    return 0

def riclassifica(cfg, mostra=None):
    """Riapplica i criteri al bacino, senza aprire Outlook. Dice cosa e'
    cambiato rispetto all'ultima volta: quali messaggi entrano ora che prima
    erano scartati, e quali escono. E' il modo per affinare i criteri in
    pochi secondi invece che in dieci minuti."""
    b = leggi_json(BACINO, None)
    if not b or not b.get("messaggi"):
        err(f"nessun bacino: lancia prima  python pec_intake.py --raccogli")
        return 1
    msg = b["messaggi"]
    criteri = cfg.get("criteri", {})
    usa_corpo = bool(cfg.get("usa_corpo", False))
    log(f"bacino del {b.get('aggiornato','?')[:16]} \u2014 {len(msg)} messaggi "
        f"({b.get('periodo','')})")

    class Finto:
        """Il messaggio salvato, con la stessa faccia di quello di Outlook,
        cosi' valuta() non sa la differenza."""
        def __init__(s, d):
            s._d = d
            s.Subject = d.get("oggetto", "")
            s.SenderEmailAddress = d.get("mittente", "")
            s.SenderName = d.get("mittente", "")
            s.Body = d.get("corpo", "")
            s.To = d.get("destinatari", "")
            s.CC = ""
            s.Attachments = type("A", (), {
                "Count": len(d.get("allegati") or []),
                "Item": staticmethod(lambda j, l=d.get("allegati") or []:
                                     type("F", (), {"FileName": l[j-1]})())})()
            s.PropertyAccessor = type("P", (), {
                "GetProperty": staticmethod(lambda k: (_ for _ in ()).throw(Exception()))})()

    presi, motivi, cambiati = 0, {}, []
    for d in msg:
        try:
            ok, motivo = valuta(Finto(d), criteri, usa_corpo)
        except Exception as e:
            ok, motivo = False, f"errore: {type(e).__name__}"
        prima = d.get("esito")
        d["esito"], d["motivo"] = ok, motivo
        if ok:
            presi += 1
        else:
            motivi[motivo] = motivi.get(motivo, 0) + 1
        if prima is not None and prima != ok:
            cambiati.append((d, prima, ok))

    scrivi_json(BACINO, b)
    log("")
    log("=" * 68)
    log(f"PRESI: {presi}    SCARTATI: {len(msg) - presi}")
    log("=" * 68)
    if cambiati:
        entrati = [x for x in cambiati if x[2]]
        usciti = [x for x in cambiati if not x[2]]
        log("")
        log(f"RISPETTO ALL'ULTIMA VOLTA: {len(entrati)} entrano, {len(usciti)} escono")
        for eti, l in (("ORA ENTRANO", entrati), ("ORA ESCONO", usciti)):
            if not l:
                continue
            log("")
            log(f"  {eti}:")
            for d, _, _ in l[:10]:
                log(f"    {d['data'][:10]}  {d['oggetto'][:60]}")
                log(f"                {d['motivo'][:60]}")
            if len(l) > 10:
                log(f"    ... e altri {len(l)-10}")
    log("")
    log("PERCHE' SONO STATI SCARTATI:")
    for m, n in sorted(motivi.items(), key=lambda x: -x[1])[:16]:
        log(f"  {n:6d}  {m}")
    if mostra:
        q = str(mostra).lower()
        l = [d for d in msg if not d["esito"] and q in (d.get("motivo") or "").lower()]
        log("")
        log(f"I PRIMI SCARTATI PER '{mostra}' ({len(l)} in tutto):")
        for d in l[:25]:
            log(f"    {d['data'][:10]}  {d['cartella'][:16]:<18}{d['oggetto'][:64]}")
    log("")
    log("Se gli scarti non convincono, correggi config_pec.json e rilancia")
    log("questo comando: risponde subito. Quando torna, lancia l'ingestione.")
    return 0

def cerca_pec(cfg, chiave):
    """Perche' quella PEC non e' diventata un fascicolo?

    Cerca nelle cartelle di Outlook i messaggi che contengono una parola —
    un cognome, un numero di ruolo — e per ognuno dice cosa il programma ne
    farebbe: se lo scarta e per quale criterio, o se lo prende e come.
    E' il modo per rispondere alla domanda senza tirare a indovinare."""
    ns = collega_outlook()
    if not ns:
        return 1
    cartelle = scegli_cartelle(ns, cfg.get("cartelle", []), cfg.get("escludi_cartelle", []))
    if not cartelle:
        err("nessuna cartella trovata")
        return 1
    criteri = cfg.get("criteri", {})
    usa_corpo = bool(cfg.get("usa_corpo", False))
    q = str(chiave).strip().lower()
    varianti = {q, q.replace("/", "-"), q.replace("-", "/")}
    m0 = re.match(r"^0*(\d+)\s*[/-]\s*(\d{2,4})$", q)
    if m0:                      # 1250/2026 vale anche come 1250/26 e 001250/26
        n, a = m0.group(1), m0.group(2)
        a2 = a[-2:]
        for x in (n, n.zfill(6)):
            for y in (a, a2):
                varianti |= {f"{x}/{y}", f"{x}-{y}", f"{x} {y}"}
    log(f"cerco: {', '.join(sorted(varianti))}")
    log("")

    trovati = esaminati = 0
    for cart in cartelle:
        try:
            items = prop(cart, "Items")
            items.Sort("[ReceivedTime]", True)
        except Exception:
            continue
        it = items.GetFirst()
        while it is not None:
            m, it = it, items.GetNext()
            if not e_un_messaggio(m):
                continue
            esaminati += 1
            if esaminati % 500 == 0:
                log(f"    ...{esaminati} messaggi esaminati")
            ogg = testo(prop(m, "Subject"))
            # qui il corpo si legge SEMPRE: e' una ricerca, non la selezione.
            # Il numero di ruolo di una notifica sta quasi sempre nel corpo o
            # dentro la busta, non nell'oggetto.
            corpo = testo(prop(m, "Body"))[:20000]
            allegati = []
            try:
                for j in range(1, m.Attachments.Count + 1):
                    allegati.append(testo(prop(m.Attachments.Item(j), "FileName")))
            except Exception:
                pass
            dove = (ogg + " " + corpo + " " + " ".join(allegati)).lower()
            visto = any(v in dove for v in varianti)
            if not visto and allegati:
                # dentro postacert.eml c'e' il messaggio vero: senza aprirlo
                # si cerca solo nella scorza esterna
                try:
                    import tempfile
                    with tempfile.TemporaryDirectory() as td:
                        for j in range(1, m.Attachments.Count + 1):
                            a = m.Attachments.Item(j)
                            n = testo(prop(a, "FileName"))
                            if not re.search(r"postacert\.eml|\.eml$", n, re.I):
                                continue
                            p = Path(td) / "b.eml"
                            a.SaveAsFile(str(p))
                            testo_b = p.read_bytes().decode("utf-8", "ignore")[:200000]
                            if any(v in testo_b.lower() for v in varianti):
                                visto = "dentro la busta"
                                break
                except Exception:
                    pass
            if not visto:
                continue
            trovati += 1
            quando = prop(m, "ReceivedTime", None)
            mit = mittente_vero(m)
            log("=" * 68)
            log(f"  {str(quando)[:16]}  {testo(prop(cart, 'Name'))}")
            log(f"  oggetto : {ogg[:88]}")
            log(f"  mittente: {mit[:60]}")
            log(f"  allegati: {len(allegati)}"
                + ("  " + ", ".join(allegati[:4]) if allegati else ""))
            try:
                ok, motivo = valuta(m, criteri, usa_corpo)
                log(f"  ESITO   : {'PRESA' if ok else 'SCARTATA'}  ({motivo})")
            except Exception as e:
                log(f"  ESITO   : errore nella valutazione: {e}")
            if visto == "dentro la busta":
                log("  il riferimento e' DENTRO LA BUSTA (postacert.eml)")
            dove_q = [v for v in varianti if v in ogg.lower()]
            if visto == "dentro la busta":
                pass
            elif dove_q:
                log(f"  il riferimento e' NELL'OGGETTO: {dove_q[0]}")
            elif any(v in " ".join(allegati).lower() for v in varianti):
                log("  il riferimento e' nel NOME DI UN ALLEGATO")
            else:
                log("  il riferimento e' solo nel CORPO del messaggio")
    log("")
    log("=" * 68)
    if trovati:
        log(f"{trovati} messaggi su {esaminati} contengono '{chiave}'.")
        log("Se sono tutti SCARTATI, il motivo indicato dice quale criterio")
        log("li esclude. Se sono PRESI, il problema e' piu' avanti.")
    else:
        log(f"NESSUN messaggio contiene '{chiave}' fra i {esaminati} esaminati.")
        log("Quella comunicazione non e' mai arrivata qui: o sta in un'altra")
        log("cartella di Outlook, o non e' mai stata inviata a questa casella.")
    return 0

def promuovi(cfg, simula=False):
    """IL PONTE FRA LE DUE CARTELLE.

    Una PEC arriva prima che il fascicolo esista: la sua cartella resta in
    IN ARRIVO. Quando il fascicolo entra dopo — da SISCO o creato a mano —
    nessuno torna indietro a prendere quella PEC, e i documenti restano
    separati dal fascicolo a cui appartengono.

    Questo comando rimette in fila le due cose: per ogni fascicolo in registro
    senza cartella, cerca la PEC che gli corrisponde e la sposta in RICORSI ES
    col nome di convenzione. Da lanciare dopo ogni caricamento SISCO."""
    base_p = cfg.get("base_json")
    radice = cfg.get("archivio_fascicoli")
    arrivo = cfg.get("destinazione")
    if not (base_p and radice and arrivo):
        err("servono base_json, archivio_fascicoli e destinazione in percorsi.json")
        return 1

    def ass(p):
        q = Path(p)
        return q if q.is_absolute() else BASE / q

    base = leggi_json(ass(base_p), None)
    if not base:
        err(f"archivio non leggibile: {ass(base_p)}")
        return 1
    rad, arr = ass(radice), ass(arrivo)
    if not arr.exists():
        err(f"non trovo {arr}")
        return 1
    rad.mkdir(parents=True, exist_ok=True)
    FASC["formato_mese"] = _impara_formato_mese(rad) or "%m%y %b"

    num = lambda s: re.sub(r"[^0-9]", "", str(s or ""))

    # --- 1. cosa c'e' gia' in RICORSI ES, per non duplicare ---
    gia = set()
    for d in rad.rglob("*"):
        if d.is_dir():
            n = _legge_nome_cartella(d.name)
            if n:
                gia.add(n)
    # ripartite per cartella di mese: cosi' si confronta con quello che vede
    # il registro, e si vede subito se qualcuno dei due si ferma prima
    per_madre = {}
    for d in rad.rglob("*"):
        if d.is_dir() and _legge_nome_cartella(d.name):
            m = d.relative_to(rad).parts[0] if len(d.relative_to(rad).parts) > 1 else "(radice)"
            per_madre[m] = per_madre.get(m, 0) + 1
    log(f"ARCHIVIO: {rad.resolve()}")
    log(f"POSTA   : {arr.resolve()}")
    log(f"{len(gia)} cartelle gia' in {rad.name}, in {len(per_madre)} cartelle di mese")
    ordinate = sorted(per_madre.items(), key=lambda x: -x[1])
    log("  " + ", ".join(f"{m} {n}" for m, n in ordinate[:14])
        + (", ..." if len(ordinate) > 14 else ""))

    # --- 2. le PEC in IN ARRIVO, indicizzate per ruolo e per SISCO ---
    per_rg, per_sisco, quante = {}, {}, 0
    for d in arr.rglob("scheda_pec.json"):
        s = leggi_json(d, None)
        if not s:
            continue
        quante += 1
        cart = d.parent
        rif = s.get("riferimenti") or {}
        if rif.get("rg"):
            per_rg.setdefault(num(rif["rg"]), []).append((cart, s))
        if rif.get("sisco"):
            per_sisco.setdefault(_compatta(rif["sisco"]), []).append((cart, s))
        # Il numero di pratica sta spesso solo nell'OGGETTO: "RICHIESTA
        # ISTRUTTORIA (Sede: X - Pratica: 533/26/CO/1 - NOME)". Senza cercarlo
        # li', i fascicoli nati da quegli oggetti non ritrovano mai la loro PEC.
        for f2 in fascicoli_dall_oggetto(s.get("oggetto") or ""):
            per_sisco.setdefault(_compatta(f2["sisco"]), []).append((cart, s))
        for m2 in re.finditer(r"\b(\d{1,4}\s*/\s*\d{2}\s*/\s*(?:CO|AO|CV)"
                              r"(?:\s*/\s*\d)?)\b", testo(s.get("oggetto")), re.I):
            per_sisco.setdefault(_compatta(m2.group(1)), []).append((cart, s))
    log(f"{quante} PEC in {arr.name}: {len(per_rg)} ruoli, {len(per_sisco)} numeri SISCO")

    # --- 3. i fascicoli che aspettano una cartella ---
    ud_per_f = {}
    for u in base.get("udienze", []):
        if u.get("data") and u.get("fascicolo_id"):
            ud_per_f.setdefault(u["fascicolo_id"], []).append(u["data"])

    spostate, uniti, senza_pec, senza_data, gia_a_posto = 0, 0, 0, 0, 0
    dall_udienza_pec = 0
    del UDIENZE_TROVATE[:]
    for f in base.get("fascicoli", []):
        if f.get("_eliminato") or f.get("stato") == "archiviato":
            continue
        nome = testo(f.get("ricorrente")).strip()
        if not nome:
            continue
        date_ud = sorted(ud_per_f.get(f["id"], []), reverse=True)

        # I fascicoli nati dagli oggetti delle PEC non hanno udienza: senza,
        # il nome della cartella non si compone e restano fermi per sempre.
        # Ma la PEC che li riguarda l'udienza spesso ce l'ha: si prende da li'.
        trovate = (per_rg.get(num(f.get("rg"))) or []) if f.get("rg") else []
        if not trovate and f.get("sisco"):
            trovate = per_sisco.get(_compatta(f["sisco"])) or []
        if not date_ud and trovate:
            for _, s in trovate:
                u = testo(s.get("data_udienza"))[:10]
                if re.match(r"^\d{4}-\d{2}-\d{2}$", u):
                    date_ud = [u]
                    dall_udienza_pec += 1
                    UDIENZE_TROVATE.append({"id": f["id"], "ricorrente": nome,
                                            "rg": f.get("rg"), "sisco": f.get("sisco"),
                                            "data": u})
                    break

        if not date_ud:
            senza_data += 1
            continue
        try:
            a, m, g = date_ud[0].split("-")
            d_ud = date(int(a), int(m), int(g))
        except Exception:
            senza_data += 1
            continue

        chiave = _chiave_fascicolo(nome, f"{d_ud:%d%m%y}")
        if chiave in gia:
            gia_a_posto += 1
            continue                       # la cartella c'e' gia'

        if not trovate:
            senza_pec += 1
            continue

        madre = rad / nome_mese(d_ud, FASC["formato_mese"])
        dest = madre / pulisci_nome(f"{nome} UD {d_ud:%d%m%y}", 120)
        if simula:
            log(f"  {f.get('rg') or f.get('sisco')}  {nome[:30]:<32}"
                f"<- {len(trovate)} PEC")
            spostate += 1
            continue
        try:
            madre.mkdir(parents=True, exist_ok=True)
            prima = trovate[0][0]
            if not dest.exists():
                prima.rename(dest)
                spostate += 1
            else:
                for x in prima.iterdir():
                    if not (dest / x.name).exists():
                        shutil.move(str(x), str(dest / x.name))
                shutil.rmtree(prima, ignore_errors=True)
                uniti += 1
            # le altre PEC dello stesso fascicolo entrano nella stessa cartella
            for altra, _ in trovate[1:]:
                if not altra.exists():
                    continue
                for x in altra.iterdir():
                    q = dest / x.name
                    if q.exists():
                        # stesso nome ma altro documento: non si sovrascrive
                        # e non si butta. Se e' identico, si lascia perdere.
                        try:
                            uguali = (x.is_file() and q.is_file()
                                      and x.stat().st_size == q.stat().st_size
                                      and x.read_bytes() == q.read_bytes())
                        except Exception:
                            uguali = False
                        if uguali:
                            continue
                        i = 2
                        while (dest / f"{x.stem} ({i}){x.suffix}").exists():
                            i += 1
                        q = dest / f"{x.stem} ({i}){x.suffix}"
                    shutil.move(str(x), str(q))
                shutil.rmtree(altra, ignore_errors=True)
                uniti += 1
            marca_cartella(dest, {"id": f["id"], "ricorrente": nome,
                                  "rg": f.get("rg"), "sisco": f.get("sisco"),
                                  "udienze": date_ud}, {})
            gia.add(chiave)
            LANCIO.setdefault("promosse", []).append(str(dest))
        except Exception as e:
            err(f"{nome}: {e}")

    log("")
    log("=" * 66)
    log(f"{'SAREBBERO SPOSTATE' if simula else 'SPOSTATE'}: {spostate} cartelle da "
        f"{arr.name} a {rad.name}")
    if uniti:
        log(f"  {uniti} PEC unite a cartelle gia' esistenti")
    if dall_udienza_pec:
        log(f"  {dall_udienza_pec} hanno preso la data di udienza dalla PEC")
        # quelle udienze vanno anche nel registro, altrimenti la cartella
        # esiste ma il fascicolo continua a risultare senza udienza
        col = ["CODICE", "GIUDICE", "MATERIA", "ARGOMENTO", "FASCICOLO",
               "DATA UDIENZA", "DATA PEC", "DIFENSORE INPS", "DATA COSTITUZIONE",
               "SISCO", "RG", "SEDE", "GRADO", "AVV. CONTROPARTE", "STATO", "NOTE"]
        fuori = BASE / "_UDIENZE_DALLE_PEC.csv"
        with open(fuori, "w", newline="", encoding="utf-8") as fh:
            fh.write("\ufeff")
            w = csv.DictWriter(fh, col, delimiter=";", quoting=csv.QUOTE_ALL)
            w.writeheader()
            for x in UDIENZE_TROVATE:
                w.writerow({c: "" for c in col} | {
                    "CODICE": x["id"], "FASCICOLO": x["ricorrente"],
                    "DATA UDIENZA": "/".join(reversed(x["data"].split("-"))),
                    "SISCO": x["sisco"] or "", "RG": x["rg"] or "",
                    "NOTE": "udienza letta dalla PEC"})
        log(f"     {fuori.name}: importalo dal registro per registrarle")
    log(f"  {gia_a_posto} fascicoli avevano gia' la loro cartella")
    log(f"  {senza_pec} fascicoli non hanno una PEC che li riguardi")
    if senza_data:
        log(f"  {senza_data} fascicoli senza udienza: il nome non e' componibile")
    log("=" * 66)
    resta = sum(1 for d in arr.rglob("scheda_pec.json"))
    if resta:
        log("")
        log(f"In {arr.name} restano {resta} PEC: riguardano fascicoli che il registro")
        log("non ha ancora. Rilancia questo comando dopo ogni caricamento SISCO.")
    if simula:
        log("Era una prova: rilancia senza --simulazione per spostarle davvero.")
    elif spostate or uniti:
        log("Ora dal registro: Cartelle e manutenzione -> Ricollega tutte.")
    return 0

def azzera(cfg, anche_fascicoli=False):
    """Riporta l'ingestione allo stato iniziale: svuota il transito e la
    memoria dei messaggi visti. Con 'tutto' toglie anche le comunicazioni
    depositate nei fascicoli. Non cancella mai i documenti che c'erano prima."""
    dest = Path(cfg.get("destinazione") or "IN ARRIVO")
    if not dest.is_absolute():
        dest = BASE / dest

    conta = {"cartelle": 0, "code": 0, "registri": 0}
    if dest.exists():
        for x in list(dest.iterdir()):
            if x.is_dir():
                if rimuovi(x):
                    conta["cartelle"] += 1
            else:
                rimuovi(x)

    if anche_fascicoli:
        radice = cfg.get("archivio_fascicoli")
        if radice:
            p = Path(radice)
            if not p.is_absolute():
                p = BASE / p
            if p.exists():
                # anche quelle messe da parte in una corsa precedente: il
                # blocco di OneDrive e' temporaneo, il nome resterebbe per sempre
                code = [x for x in p.rglob("*")
                        if x.is_dir() and (x.name == "COMUNICAZIONI"
                                           or x.name.startswith("_DA CANCELLARE"))]
                for coda in code:
                    if rimuovi(coda):
                        conta["code"] += 1
                for reg in list(p.rglob("comunicazioni.json")):
                    if rimuovi(reg):
                        conta["registri"] += 1
                # le cartelle di fascicolo create dall'ingestione: si riconoscono
                # dal marcatore. Si tolgono solo se non contengono altro, cosi'
                # una cartella poi usata da qualcuno resta intatta.
                conta["fascicoli"] = 0
                for mk in list(p.rglob(".fascicolo.json")):
                    d = mk.parent
                    try:
                        resto = [x for x in d.iterdir() if x.name != ".fascicolo.json"]
                        if resto:
                            continue
                        if rimuovi(d):
                            conta["fascicoli"] += 1
                        if d.parent.exists() and not any(d.parent.iterdir()):
                            d.parent.rmdir()
                    except Exception as e:
                        err(f"non rimossa {d.name}: {e}")

    for f in (ARCHIVIO, LANCI):
        try:
            if f.exists():
                f.unlink()
        except Exception as e:
            err(f"non rimosso {f.name}: {e}")

    rimaste = []
    for radice in (dest, Path(cfg.get("archivio_fascicoli") or "")):
        try:
            if radice and str(radice) and radice.exists():
                rimaste += [str(x) for x in radice.rglob("_DA CANCELLARE*")]
        except Exception:
            pass

    log("")
    log("AZZERATO:")
    log(f"  {conta['cartelle']} cartelle di transito rimosse da {dest.name}")
    if anche_fascicoli:
        log(f"  {conta['code']} sottocartelle COMUNICAZIONI rimosse dai fascicoli")
        log(f"  {conta['registri']} registri comunicazioni.json rimossi")
        log(f"  {conta.get('fascicoli', 0)} cartelle di fascicolo create dall'ingestione e "
            "rimaste vuote")
    else:
        log("  le comunicazioni depositate nei fascicoli NON sono state toccate")
        log("  (per toglierle anche quelle:  --azzera tutto)")
    log("  memoria dei messaggi visti e storico dei lanci cancellati")
    log("")
    if rimaste:
        log("")
        log(f"ATTENZIONE: {len(rimaste)} cartelle sono bloccate da OneDrive e sono state")
        log("messe da parte col prefisso '_DA CANCELLARE'.")
        log("")
        log("Metti OneDrive in pausa (clic sull'icona nella barra > Sospendi la")
        log("sincronizzazione) e rilancia --azzera: verranno riprovate.")
        log("In alternativa cancellale a mano quando la sincronizzazione ha finito.")
        for x in rimaste[:8]:
            log("   " + x)
    log("")
    log("Le cartelle dei fascicoli e i loro documenti non sono stati toccati.")
    return 0

def pulisci_doppioni(cfg):
    """Toglie cio' che una rilettura ha duplicato: righe ripetute nelle code dei
    fascicoli e cartelle di transito con il vecchio nome rimaste accanto a
    quelle nuove. Non tocca i documenti."""
    tolte_righe = tolte_cartelle = 0

    radice = cfg.get("archivio_fascicoli")
    if radice:
        p = Path(radice)
        if not p.is_absolute():
            p = BASE / p
        for f in p.rglob("comunicazioni.json") if p.exists() else []:
            reg = leggi_json(f, {})
            elenco = reg.get("comunicazioni") or []
            visti, puliti, rifatte = set(), [], 0
            for v in elenco:
                if not chiave_vera(v.get("chiave")):
                    v["chiave_scritta"] = v.get("chiave")
                    v["chiave"] = chiave_comunicazione(v, f.parent.name)
                    rifatte += 1
                k = v["chiave"]
                if k in visti:
                    continue
                visti.add(k)
                puliti.append(v)
            if rifatte:
                log(f"  {f.parent.name}: {rifatte} chiavi finte sostituite")
            if len(puliti) != len(elenco) or rifatte:
                tolte_righe += len(elenco) - len(puliti)
                reg["comunicazioni"] = puliti
                scrivi_json(f, reg)
                log(f"  {f.parent.name}: {len(elenco)-len(puliti)} righe ripetute rimosse")

    dest = Path(cfg.get("destinazione") or "IN ARRIVO")
    if not dest.is_absolute():
        dest = BASE / dest
    if dest.exists():
        # una cartella e' superata se un'altra porta la stessa chiave-messaggio
        per_chiave = {}
        for s in dest.rglob("scheda_pec.json"):
            d = leggi_json(s, {})
            k = d.get("chiave")
            if not k:
                continue
            per_chiave.setdefault(k, []).append((s.parent, d))
        for k, gruppo in per_chiave.items():
            if len(gruppo) < 2:
                continue
            # tengo quella con il nome piu' informativo: chi ha il ricorrente vince
            gruppo.sort(key=lambda g: (0 if g[1].get("ricorrente") else 1,
                                       -len(g[1].get("documenti") or [])))
            tenuta = gruppo[0][0]
            for vecchia, _ in gruppo[1:]:
                try:
                    shutil.rmtree(vecchia)
                    tolte_cartelle += 1
                    log(f"  rimossa {vecchia.name} (superata da {tenuta.name})")
                except Exception as e:
                    err(f"non rimossa {vecchia.name}: {e}")

    log("")
    log(f"PULIZIA COMPLETATA: {tolte_righe} righe ripetute, {tolte_cartelle} cartelle superate")
    log("I documenti non sono stati toccati.")
    return 0

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ricognizione", action="store_true", help="legge e basta, non scrive")
    ap.add_argument("--rileggi", action="store_true", help="ignora l'archivio e rielabora tutto")
    ap.add_argument("--dal", help="data iniziale AAAA-MM-GG (ha la precedenza su giorni_indietro)")
    ap.add_argument("--al", help="data finale AAAA-MM-GG")
    ap.add_argument("--annulla", nargs="?", const="ultimo",
                    help="annulla un lancio: --annulla ultimo, oppure --annulla 20260902_205000")
    ap.add_argument("--lanci", action="store_true", help="elenca i lanci registrati")
    ap.add_argument("--destinazione", nargs="+",
                    help="cartella dove depositare, invece di quella "
                                          "in percorsi.json (per le prove)")
    ap.add_argument("--simulazione", action="store_true",
                    help="prova a vuoto: ignora il registro e l'archivio dei fascicoli, "
                         "costruisce tutto da zero nella cartella indicata")
    ap.add_argument("--raccogli", action="store_true",
                    help="legge Outlook UNA VOLTA e salva il bacino in locale: "
                         "da li' si riclassifica in un istante")
    ap.add_argument("--riclassifica", action="store_true",
                    help="riapplica i criteri al bacino, senza aprire Outlook")
    ap.add_argument("--mostra", metavar="MOTIVO",
                    help="con --riclassifica: elenca gli scartati per un motivo")
    ap.add_argument("--dagli-oggetti", action="store_true",
                    help="raccoglie i fascicoli nominati negli oggetti delle PEC "
                         "(elenchi di ricorsi, moduli di richiesta istruttoria)")
    ap.add_argument("--cerca", metavar="TESTO",
                    help="cerca nelle PEC un cognome o un numero di ruolo e dice "
                         "per ognuna se verrebbe presa o scartata, e perche'")
    ap.add_argument("--promuovi", action="store_true",
                    help="sposta in RICORSI ES le PEC dei fascicoli che nel "
                         "frattempo sono entrati in registro")
    # nargs="+" e poi si riuniscono i pezzi: i percorsi di Windows hanno gli
    # spazi - "OneDrive - INPS\SISCO AGENT\APPO" - e senza virgolette il
    # prompt li spezza in piu' argomenti. Argparse rispondeva
    # "unrecognized arguments: - INPS\SISCO AGENT\APPO", che non dice a
    # nessuno di mettere le virgolette.
    ap.add_argument("--converti", metavar="PERCORSO", nargs="+",
                    help="converte le estrazioni SISCO da .xls a .xlsx "
                         "(un file o una cartella intera). Se il percorso ha "
                         "spazi funziona lo stesso, con o senza virgolette")
    ap.add_argument("--leggi-cartelle", action="store_true",
                    help="legge i documenti dei fascicoli in archivio e propone i dati mancanti")
    ap.add_argument("--azzera", nargs="?", const="transito",
                    help="svuota IN ARRIVO e la memoria; '--azzera tutto' toglie anche "
                         "le comunicazioni depositate nei fascicoli")
    ap.add_argument("--ogni", type=int, metavar="MINUTI",
                    help="resta in ascolto e rilegge la posta ogni MINUTI minuti "
                         "(es. --ogni 5). Si ferma con Ctrl+C")
    ap.add_argument("--pulisci", action="store_true",
                    help="rimuove i doppioni gia' formati e si ferma, senza leggere Outlook")
    ap.add_argument("--config", default=str(CONFIG))
    args = ap.parse_args()

    # I percorsi spezzati dal prompt tornano interi, una volta sola e qui:
    # da qui in avanti sono stringhe come sempre, e nessun altro punto del
    # programma deve sapere che sono arrivati a pezzi.
    for _campo in ("converti", "destinazione"):
        _v = getattr(args, _campo, None)
        if isinstance(_v, list):
            setattr(args, _campo, " ".join(_v) if _v else None)

    cfg_path = Path(args.config)
    if not cfg_path.exists():
        # I predefiniti leggono TUTTE le cartelle e non hanno i criteri messi a
        # punto: creare il file e proseguire significa lavorare a vuoto per
        # mezz'ora. Meglio fermarsi e dirlo.
        scrivi_json(cfg_path, CONFIG_DEFAULT)
        err("=" * 70)
        err(f"{cfg_path.name} non c'era: ne ho creato uno con i criteri predefiniti.")
        err("NON e' quello giusto: legge tutte le cartelle della casella e non ha")
        err("i criteri messi a punto per il contenzioso.")
        err("")
        err("Rimetti il tuo config_pec.json in questa cartella e rilancia.")
        err("=" * 70)
        return 1
    cfg = leggi_json(cfg_path, None)
    if cfg is None:
        err("=" * 70)
        err(f"CONFIGURAZIONE NON LETTA: {cfg_path.name}")
        # mostrare la riga com'e' scritta risolve nove volte su dieci
        try:
            import json as _j
            testo_cfg = cfg_path.read_text(encoding="utf-8")
            try:
                _j.loads(testo_cfg)
            except _j.JSONDecodeError as e:
                righe = testo_cfg.split("\n")
                err("")
                err(f"  {e.msg}, riga {e.lineno} colonna {e.colno}")
                err("")
                for i in range(max(0, e.lineno - 3), min(len(righe), e.lineno + 2)):
                    segno = ">>" if i == e.lineno - 1 else "  "
                    err(f"  {segno} {i+1:4d} | {righe[i][:100]}")
                err("")
                if "delimiter" in e.msg:
                    err("  Manca una virgola fra due voci, oppure i due punti dopo il nome.")
                elif "Expecting" in e.msg and "value" in e.msg:
                    err("  C'e' una virgola di troppo prima di una parentesi che chiude.")
                elif "Unterminated" in e.msg:
                    err("  Ci sono virgolette aperte e mai chiuse.")
        except Exception:
            pass
        err("")
        err("  Senza configurazione userei i criteri predefiniti: tutte le cartelle")
        err("  della casella e nessuno dei tuoi criteri. Mi fermo.")
        err("  Se non riesci a correggerlo, riscarica config_pec.json e rimetti")
        err("  i tuoi giudici in 'giudici_seguiti'.")
        err("=" * 70)
        return 1
    mancanti = [k for k in ("cartelle", "criteri", "tipi_pec") if k not in cfg]
    if mancanti:
        err(f"CONFIGURAZIONE INCOMPLETA: mancano {', '.join(mancanti)} in {cfg_path.name}")
        err("Probabilmente e' un file vecchio o troncato. Mi fermo.")
        return 1
    log(f"programma: {VERSIONE}")
    log(f"configurazione: {cfg_path.name} \u2014 {len(cfg.get('cartelle') or [])} cartelle, "
        f"{len(cfg.get('criteri', {}).get('oggetto_forte') or [])} criteri forti")

    # --- percorsi: file separato, cosi' gli aggiornamenti dei criteri non li toccano ---
    if not PERCORSI.exists():
        scrivi_json(PERCORSI, PERCORSI_DEFAULT)
        log(f"creato {PERCORSI.name}: qui vanno i percorsi di questa installazione")
    perc = leggi_json(PERCORSI, {})
    CHIAVI_PERCORSO = ("destinazione", "archivio_fascicoli", "indice_cartelle", "base_json")
    for k in CHIAVI_PERCORSO:
        # quello che si chiede da riga di comando vince su tutto: e' il modo
        # per fare una prova senza toccare l'installazione
        if k == "destinazione" and getattr(args, "destinazione", None):
            cfg[k] = args.destinazione
            continue
        if perc.get(k):
            cfg[k] = perc[k]                      # percorsi.json ha la precedenza
        elif cfg.get(k) and not perc.get(k):
            perc[k] = cfg[k]                      # migro quelli gia' scritti nel vecchio file
    if not any(cfg.get(k) for k in ("archivio_fascicoli", "base_json", "indice_cartelle")):
        auto = cerca_percorsi()
        if auto:
            log("percorsi non indicati: trovati da solo ->")
            for k, v in auto.items():
                cfg[k] = v
                perc.setdefault(k, v)
                log(f"    {k}: {v}")
    scrivi_json(PERCORSI, {**PERCORSI_DEFAULT, **{k: v for k, v in perc.items() if v}})
    sposta_archivio_accanto_al_registro(cfg, ricognizione=bool(args.ricognizione))

    global CFG
    CFG = cfg
    criteri = cfg.get("criteri", {})
    global TIPI
    TIPI = cfg.get("tipi_pec", {})
    global CASELLE
    CASELLE = cfg.get("caselle_notifica", {})
    global FASI
    FASI = cfg.get("fasi_ctu", {})
    dest_radice = Path(cfg.get("destinazione", "IN ARRIVO"))
    if not dest_radice.is_absolute():
        dest_radice = BASE / dest_radice

    archivio = leggi_json(ARCHIVIO, {"visti": {}})
    # in simulazione si riparte da zero anche sulla memoria dei messaggi:
    # altrimenti si guarda un campione casuale invece del periodo scelto
    visti_mail = {} if (args.rileggi or getattr(args, "simulazione", False)) \
                 else archivio.get("visti", {})   # vedi da_capo, piu' sotto
    global COPIE, AVVISI, ACCANTONATI, SOLO_CON_DOCUMENTI, LAVORATE, IN_CODA
    LAVORATE = []
    IN_CODA = []
    # Da zero significa DA ZERO: anche la memoria delle copie e degli avvisi.
    # Tenerla faceva risultare copia di un lancio precedente ogni messaggio.
    da_capo = args.rileggi or getattr(args, "simulazione", False)


    COPIE = {} if da_capo else archivio.get("copie", {})
    FIRME_LANCIO.clear()
    AVVISI = [] if da_capo else archivio.get("avvisi", [])
    ACCANTONATI = [] if da_capo else archivio.get("accantonati", [])
    SOLO_CON_DOCUMENTI = bool(cfg.get("solo_con_documenti", True))
    global ESECUZIONI
    ESECUZIONI = bool(cfg.get("accantona_esecuzioni", True))
    global SCARTA_INTERNE, ERRORI, SIMULAZIONE
    ERRORI = []
    SCARTA_INTERNE = bool(cfg.get("scarta_comunicazioni_interne", True))
    global NOMI_INTERNI
    NOMI_INTERNI = set(cfg.get("nomi_interni") or [])
    global SCARTA_INOLTRI, IDENTIFICATIVI
    SCARTA_INOLTRI = bool(cfg.get("scarta_inoltri_interni", True))
    global ACCANTONA
    ACCANTONA = bool(cfg.get("accantona_non_di_competenza", True))
    IDENTIFICATIVI = {}

    try:
        import olefile  # noqa: F401
    except ImportError:
        err("libreria olefile assente: le PEC incorporate da Outlook come .msg")
        err("non potranno essere aperte. Installala con:  pip install olefile")

    # collaudo a vuoto: meglio fermarsi ora che scoprire a meta' corsa che
    # una funzione manca e contare migliaia di "elementi illeggibili"
    def _finto(oggetto):
        class F:
            Class = 43
            Subject = oggetto
            SenderEmailAddress = "avv@pec.it"; SenderName = "avv"; To = ""; CC = ""
            class Attachments:
                Count = 0
                def Item(self, j): return None
        return F()

    # due prove: la prima scatta su una parola, la seconda SOLO su un modello
    # regolare, cosi' viene percorso tutto il riconoscimento
    prove = [("POSTA CERTIFICATA: notificazione ai sensi della legge 53", True),
             ("Sentenza n. 1219/25 La Mesa Giancarlo", True),
             ("Conferma pagamento bonifico periodico", False)]
    try:
        for oggetto, atteso in prove:
            ok_prova, motivo_prova = valuta(_finto(oggetto), criteri, False)
            if ok_prova != atteso:
                err(f"ATTENZIONE: '{oggetto[:44]}' doveva essere "
                    f"{'preso' if atteso else 'scartato'} e invece "
                    f"{'e stato preso' if ok_prova else 'e stato scartato'} ({motivo_prova}). "
                    "Controlla config_pec.json.")
        oggetto_reale(prove[0][0])
        tipo_di_pec("notificazione ai sensi", [], TIPI)
        riferimenti("RG 1250/2026")
        fase_ctu("relazione peritale", [], FASI)
        casella_di_notifica(_finto("x"), "", CASELLE, None, None)
        leggi_comunicazione_tribunale(
            "Comunicazione di cancelleriaNumero di Ruolo generale: 1/2026"
            "Ricorr. principale: I.N.P.S. Resist. principale: ROSSI MARIOOggetto: X")
        nome_da_pratica("(Sede: SIRACUSA - Pratica: 1/26/CO/1 - ROSSI MARIO)")

        # elaborazione a vuoto: percorre TUTTO il flusso, compresi i rami che
        # si attivano solo per gli avvisi senza allegati. E' il controllo che
        # mancava quando 152 messaggi su 194 finivano fra gli illeggibili.
        import tempfile
        with tempfile.TemporaryDirectory() as tmpd:
            vera = ARCHIVIO
            globals()["ARCHIVIO"] = Path(tmpd) / "prova.json"
            try:
                for con_allegati in (False,):
                    elabora_messaggio(_finto("POSTA CERTIFICATA: notificazione ai sensi legge 53"),
                                      type("C", (), {"Name": "prova", "FolderPath": "prova"})(),
                                      "prova:collaudo", "collaudo",
                                      argparse.Namespace(ricognizione=False),
                                      Path(tmpd) / "IN ARRIVO", {}, [])
            finally:
                globals()["ARCHIVIO"] = vera
                AVVISI[:] = [a for a in AVVISI if a.get("chiave") != "prova:collaudo"]
                LAVORATE[:] = [x for x in LAVORATE if x.get("chiave") != "prova:collaudo"]
                COPIE.pop(next((k for k, v in COPIE.items() if v == "__avviso__"), "-"), None)
    except Exception as e:
        err("=" * 68)
        import traceback
        err(f"GUASTO NEL CODICE: {type(e).__name__}: {e}")
        for riga in traceback.format_exc().strip().split("\n")[-4:]:
            err("  " + riga)
        err("Il riconoscimento non funziona: ogni messaggio finirebbe fra gli")
        err("'elementi illeggibili'. Corri corretto prima di elaborare.")
        err("=" * 68)
        return 1

    if args.destinazione:
        cfg["destinazione"] = args.destinazione   # gia' applicato sopra, per sicurezza
    if args.simulazione and getattr(args, "promuovi", False):
        # per --promuovi "simulazione" vuol dire solo: elenca senza spostare.
        # Non serve una destinazione nuova, e i percorsi veri servono tutti.
        SIMULAZIONE = True
    elif args.simulazione:
        SIMULAZIONE = True
        # niente registro, niente archivio: si vede cosa il sistema sa
        # ricostruire dalle sole PEC. E nulla puo' finire nei fascicoli veri.
        cfg["indice_cartelle"] = None
        cfg["base_json"] = None
        cfg["archivio_fascicoli"] = None
        cfg["puo_creare_fascicoli"] = False
        if not args.destinazione:
            err("--simulazione richiede --destinazione: indica una cartella NUOVA")
            err("  esempio: python pec_intake.py --simulazione --destinazione PROVA")
            return 1
        d = Path(cfg["destinazione"])
        if not d.is_absolute():
            d = BASE / d
        if d.exists() and any(d.iterdir()):
            err(f"la cartella {d} non e' vuota: per una prova pulita indicane una nuova")
            return 1
        log("=" * 70)
        log("SIMULAZIONE: il registro e l'archivio dei fascicoli sono ignorati.")
        log(f"Tutto viene costruito da zero in {d}")
        log("Nessun fascicolo vero verra' toccato.")
        log("=" * 70)

    if args.dal:
        cfg["dal"] = args.dal
    if args.al:
        cfg["al"] = args.al

    if args.azzera:
        tutto = str(args.azzera).lower().startswith("tut")
        print()
        print("  Sto per azzerare l'ingestione delle PEC.")
        print("  - IN ARRIVO verra' svuotata")
        print("  - la memoria dei messaggi gia' visti verra' cancellata")
        if tutto:
            print("  - le sottocartelle COMUNICAZIONI nei fascicoli verranno rimosse")
        print("  Le cartelle dei fascicoli e i loro documenti NON vengono toccati.")
        print()
        if input("  Scrivi AZZERA per confermare: ").strip().upper() != "AZZERA":
            log("annullato: non e' stato cambiato nulla")
            return 0
        return azzera(cfg, tutto)

    if getattr(args, "raccogli", False):
        return raccogli_bacino(cfg, args)

    if getattr(args, "riclassifica", False):
        return riclassifica(cfg, getattr(args, "mostra", None))

    if getattr(args, "dagli_oggetti", False):
        return raccogli_dagli_oggetti(cfg, args)

    if getattr(args, "cerca", None):
        return cerca_pec(cfg, args.cerca)

    if getattr(args, "promuovi", False):
        return promuovi(cfg, simula=bool(getattr(args, "simulazione", False)))

    if getattr(args, "converti", None):
        return converti_xls(args.converti)

    if getattr(args, "leggi_cartelle", False):
        return leggi_cartelle(cfg)

    if args.lanci:
        storico = leggi_json(LANCI, {"lanci": []})
        if not storico.get("lanci"):
            log("nessun lancio registrato")
            return 0
        log("LANCI REGISTRATI (dal piu' recente):")
        for x in reversed(storico["lanci"]):
            log(f"  {x['id']}  {x.get('iniziato','')[:16]}  {x.get('finestra','')}")
            log(f"     {len(x.get('cartelle_transito',[]))} cartelle, "
                f"{len(x.get('fascicoli_creati',[]))} fascicoli creati, "
                f"{len(x.get('code',[]))} accodate, {len(x.get('chiavi',[]))} messaggi")
        return 0

    if args.annulla:
        return annulla_lancio(args.annulla)

    if args.pulisci:
        return pulisci_doppioni(cfg)

    # ---- GLI SBARRAMENTI SULLA FINESTRA ----
    # Stanno QUI e non dove si legge l'archivio: li' passavano prima dei
    # comandi che non leggono posta, e '--azzera tutto' - che la finestra
    # non la usa nemmeno - si vedeva rifiutare con un messaggio che non
    # c'entrava niente. Da questo punto in avanti si legge la casella
    # davvero, e la finestra conta.
    # --rileggi AZZERA gli avvisi e li ricostruisce SOLO sulla finestra scelta.
    # Con giorni_indietro a 10 significa passare da 3.842 avvisi a 278, e
    # perdere per strada le 980 "da prelevare dal portale", che sono la coda di
    # lavoro piu' grande del cruscotto. Nessun messaggio lo diceva: l'archivio
    # si rimpiccioliva in silenzio e il referto mostrava numeri sensati.
    # Quindi: se in archivio ci sono avvisi piu' vecchi dell'inizio della
    # finestra, il giro non parte. Non e' un dato che si ricostruisce da solo -
    # dipende da quanto indietro va la casella - e cancellarlo per sbaglio
    # costa un pomeriggio.
    if args.rileggi and not getattr(args, "simulazione", False) \
       and not args.ricognizione:
        inizio = str(cfg.get("dal") or "")[:10]
        if not inizio:
            g = int(cfg.get("giorni_indietro") or 30)
            inizio = (datetime.now() - timedelta(days=g)).strftime("%Y-%m-%d")
        prima = [str(x.get("ricevuta") or "")[:10]
                 for x in (archivio.get("avvisi") or []) + (archivio.get("accantonati") or [])]
        prima = [d for d in prima if d and d < inizio]
        if prima:
            err("=" * 70)
            err("ANNULLATO: --rileggi rifarebbe l'archivio da zero, ma la finestra")
            err(f"e' troppo corta. In archivio ci sono {len(prima)} avvisi piu' vecchi")
            err(f"del {inizio}, il piu' vecchio del {min(prima)}: con questa finestra")
            err("sparirebbero, comprese le 'da prelevare dal portale'.")
            err("")
            err("Se vuoi rifare l'archivio, dillo per intero:")
            err(f"  python pec_intake.py --rileggi --dal {min(prima)}")
            err("Se invece volevi solo riprendere la posta recente, --rileggi non")
            err("serve: senza di lui i messaggi nuovi entrano e i vecchi restano.")
            err("=" * 70)
            return 1

    # L'ALTRA META' DELLO STESSO PERICOLO. Lo sbarramento qui sopra confronta
    # la finestra con quello che c'e' in archivio: dopo '--azzera tutto'
    # l'archivio non c'e' piu', non c'e' niente da confrontare, e un giro senza
    # --dal ricostruirebbe tutto sui soli ultimi 10 giorni. E' esattamente il
    # momento della ripartenza, quando l'archivio va rifatto per intero.
    # Quindi: archivio vuoto, nessun --dal e si sta per scrivere -> si chiede
    # la data. Un giro quotidiano non ci passa mai, perche' l'archivio e' pieno.
    if (not archivio.get("avvisi") and not archivio.get("visti")
            and not cfg.get("dal")
            and not getattr(args, "simulazione", False) and not args.ricognizione):
        g = int(cfg.get("giorni_indietro") or 30)
        err("=" * 70)
        err("ANNULLATO: l'archivio delle PEC e' vuoto e non hai detto da quando")
        err("partire, quindi leggerei solo gli ultimi " + str(g) + " giorni e l'archivio")
        err("nascerebbe gia' monco: gli avvisi piu' vecchi non entrerebbero mai.")
        err("")
        err("Se stai ripartendo da zero, di' da quando:")
        err("  python pec_intake.py --dal 2025-09-01")
        err("Metti la data piu' indietro che la casella copre: quello che non")
        err("entra adesso non entra piu'.")
        err("=" * 70)
        return 1

    apri_lancio(cfg, args)
    costruisci_indice_fascicoli(cfg)

    ns = collega_outlook()
    if not ns:
        return 1

    cartelle = scegli_cartelle(ns, cfg.get("cartelle", []), cfg.get("escludi_cartelle", []))
    if not cartelle:
        err("nessuna cartella trovata con i nomi indicati in config_pec.json")
        return 1
    log(f"{len(cartelle)} cartelle da esaminare")

    del PER_CARTELLA[:]
    filtro, descrizione_finestra = filtro_date(cfg)
    log(f"finestra: {descrizione_finestra}")

    tot = presi = saltati = gia = errori = 0
    copie_viste = 0
    avvisi_soli = 0
    da_parte = 0        # posta interna senza allegati: messa da parte, non persa
    unite = 0
    accodate = 0
    interne = 0
    inoltri = 0
    usa_corpo = bool(cfg.get("usa_corpo", True))
    motivi_no = {}
    referto = []
    sfiorati = []

    for cart in cartelle:
        # Il nome si legge UNA VOLTA, qui, e si riusa. Leggerlo dentro un
        # gestore d'errore e' stato un disastro: se l'oggetto COM non risponde,
        # la riga che doveva SEGNALARE il problema ne solleva un altro, e la
        # cartella risulta persa. Su nove cartelle ne sono cadute otto.
        try:
            nome_cart = testo(prop(cart, "Name"))
            via_cart = testo(prop(cart, "FolderPath")) or nome_cart
        except Exception:
            nome_cart = via_cart = "(cartella senza nome)"

        try:
            items = cart.Items
            # Su una cartella grande - PEC DIREZIONE ne ha 22.000 - ordinare
            # PRIMA e filtrare DOPO fa cadere Outlook: l'ordinamento tocca
            # tutti i messaggi, il filtro solo quelli del periodo. Quindi
            # prima si restringe, poi si ordina, e se anche cosi' non regge
            # si prosegue senza ordinamento invece di perdere la cartella.
            try:
                items = items.Restrict(filtro)
            except Exception as e:
                err(f"  filtro non applicato su {nome_cart}: {e}")
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception as e:
                err(f"  ordinamento non riuscito su {nome_cart}, proseguo senza: {e}")
            # gli indici cambiano se qualcosa si muove: prendo prima i riferimenti
            elenco = []
            it = items.GetFirst()
            while it is not None:
                elenco.append(it)
                it = items.GetNext()
        except Exception as e:
            # NON si salta in silenzio. Una cartella che non si apre spariva
            # dall'elenco del referto come se non esistesse: PEC DIREZIONE,
            # 22.000 messaggi, e' rimasta fuori da un giro intero senza che
            # nulla lo dicesse. L'errore va nel referto, non solo nel log.
            err(f"cartella {nome_cart}: {e}")
            PER_CARTELLA.append((via_cart + f"   <-- NON LETTA: {str(e)[:70]}", -1))
            continue

        # anche zero va detto: una cartella muta e' un'informazione, non silenzio
        PER_CARTELLA.append((via_cart, len(elenco)))
        log(f"  {via_cart}  ({len(elenco)} messaggi"
            + (", nessuno nel periodo" if not elenco else "") + ")")

        for idx, m in enumerate(elenco, 1):
            if not e_un_messaggio(m):
                continue
            tot += 1
            if idx % 250 == 0:
                log(f"    ...{idx}/{len(elenco)}")
            try:
                ok, motivo = valuta(m, criteri, usa_corpo)
            except Exception as e:
                errori += 1
                motivi_no["elemento illeggibile"] = motivi_no.get("elemento illeggibile", 0) + 1
                ERRORI.append({"tipo": type(e).__name__, "messaggio": str(e)[:120],
                               "punto": "lettura del messaggio",
                               "oggetto": testo(prop(m, "Subject"))[:70],
                               "cartella": testo(prop(cart, "Name"))})
                if errori <= 3:
                    err(f"lettura fallita ({type(e).__name__}: {e}) su "
                        f"'{testo(prop(m, 'Subject'))[:50]}'")
                continue
            if not ok:
                saltati += 1
                motivi_no[motivo] = motivi_no.get(motivo, 0) + 1
                if motivo.startswith(("un solo segnale", "solo segnali tecnici")) and len(sfiorati) < 40:
                    sfiorati.append({"data": str(prop(m, "ReceivedTime"))[:16],
                                     "oggetto": oggetto_reale(testo(prop(m, "Subject")))[:88],
                                     "motivo": motivo,
                                     "cartella": testo(prop(cart, "Name"))})
                continue
            try:
                k = chiave_stabile(m)
            except Exception as e:
                # anche questo va nell'elenco: contarlo e basta significa
                # vedere 155 illeggibili senza sapere perche'
                errori += 1
                motivi_no["elemento illeggibile"] = motivi_no.get("elemento illeggibile", 0) + 1
                ERRORI.append({"tipo": type(e).__name__, "messaggio": str(e)[:120],
                               "punto": "chiave_stabile: identificazione del messaggio",
                               "oggetto": testo(prop(m, "Subject"))[:70],
                               "cartella": testo(prop(cart, "Name"))})
                continue
            if k in visti_mail:
                # Con --rileggi la memoria parte vuota, quindi qui non ci
                # finisce mai un messaggio di un giro precedente: sono le PEC
                # presenti in DUE cartelle di Outlook, incontrate due volte in
                # questo stesso giro. L'etichetta "gia' in archivio" faceva
                # credere che il rileggi non fosse partito.
                gia += 1
                continue

            try:
                esito_el = elabora_messaggio(m, cart, k, motivo, args, dest_radice,
                                             visti_mail, referto)
                if esito_el == "copia":
                    copie_viste += 1
                elif esito_el == "avviso":
                    avvisi_soli += 1
                elif esito_el == "accantonato":
                    da_parte += 1
                elif esito_el == "unita":
                    unite += 1
                elif esito_el == "coda":
                    accodate += 1
                elif esito_el == "inoltro":
                    inoltri += 1
                    presi -= 1
                    motivi_no["inoltro interno fra colleghi"] = \
                        motivi_no.get("inoltro interno fra colleghi", 0) + 1
                elif esito_el == "interna":
                    interne += 1
                    presi -= 1
                    motivi_no["comunicazione interna in risposta a PEC"] = \
                        motivi_no.get("comunicazione interna in risposta a PEC", 0) + 1
                presi += 1
            except Exception as e:
                errori += 1
                import traceback
                tb = traceback.format_exc().strip().split("\n")
                punto = next((r.strip() for r in reversed(tb) if "pec_intake.py" in r), "")
                ERRORI.append({"tipo": type(e).__name__, "messaggio": str(e)[:120],
                               "punto": punto[:120],
                               "oggetto": testo(prop(m, "Subject"))[:70],
                               "cartella": testo(prop(cart, "Name"))})
                if errori <= 3:
                    err(f"messaggio saltato: {type(e).__name__}: {e}")
                    err("  " + punto)
                elif errori == 4:
                    err("(gli altri sono elencati per esteso nel referto)")
            continue

    if tot > 200 and presi == 0:
        err("=" * 70)
        err(f"NESSUN MESSAGGIO SELEZIONATO su {tot} letti.")
        err("Non e' normale: controlla che config_pec.json sia quello giusto e che")
        err("la voce 'cartelle' elenchi le cartelle del contenzioso.")
        err("=" * 70)

    if not args.ricognizione:
        try:
            dest_radice.mkdir(parents=True, exist_ok=True)
            salva_archivio({"visti": visti_mail, "copie": COPIE, "avvisi": AVVISI, "accantonati": ACCANTONATI,
                                   "aggiornato": datetime.now().isoformat(timespec="seconds")})
            scrivi_indice(dest_radice, LAVORATE)
        except Exception as e:
            # Un errore qui lasciava l'indice a meta' con una riga sola nel
            # log. Ora dice anche il punto: senza, si cerca per giorni.
            import traceback
            err(f"scrittura indice: {e}")
            for r in traceback.format_exc().strip().split("\n")[-3:]:
                err("    " + r.strip())
        chiudi_lancio({"elaborati": presi, "cartelle": len(LAVORATE),
                       "in_coda": len(IN_CODA), "avvisi": len(AVVISI)})

    # ---------------------------------------------------------------- referto
    r = [f"INGESTIONE PEC - {datetime.now():%d/%m/%Y %H:%M}",
         "=" * 70,
         f"programma .................. {VERSIONE}",
         "modalita' .................. " +
         ("RICOGNIZIONE (nessuna scrittura)" if args.ricognizione
          else "SIMULAZIONE DA ZERO (registro ignorato)"
          if getattr(args, "simulazione", False) else "ELABORAZIONE"),
         f"configurazione ............. {cfg_path.name}"
           + (f"  ({len(cfg.get('cartelle') or [])} cartelle indicate)"
              if cfg.get("cartelle") else "  (nessuna cartella indicata: le legge TUTTE)"),
         f"periodo .................... {descrizione_finestra}",
         ("Outlook ..................... " +
          ("l'ho avviato io, lo chiudo alla fine"
           if AVVIATO_DA_NOI["si"] else "era gia' aperto, non lo tocco")),
         f"cartelle esaminate ......... {len(cartelle)}"
           + ("   <-- ATTENZIONE: nessun messaggio letto. O il periodo e' vuoto, "
              "o Outlook non ha risposto" if not tot else ""),
         *([("    NON LETTA  " + p) if n < 0 else f"    {n:5d}  {p}"
            for p, n in PER_CARTELLA] if PER_CARTELLA else []),
         *([f"  ATTENZIONE: {sum(1 for _, n in PER_CARTELLA if n < 0)} cartelle "
            "non si sono aperte. Quello che c'e' dentro NON e' stato letto: "
            "i numeri qui sotto sono parziali."]
           if any(n < 0 for _, n in PER_CARTELLA) else []),
         f"messaggi letti ............. {tot}",
         f"corrispondenti ai criteri .. {presi + gia}",
         (f"  gia' viste in questo giro . {gia}"
          "   (la stessa PEC in due cartelle di Outlook)"
          if (getattr(args, "rileggi", False) or getattr(args, "simulazione", False))
          else f"  gia' in archivio ......... {gia}"),
         f"  {'da elaborare' if args.ricognizione else 'elaborati ora'} ............ {presi}",
         *([] if args.ricognizione else [
             f"  cartelle create ......... {len(LAVORATE)}",
             f"  soli avvisi (no cartella) {avvisi_soli}",
         (f"     di cui finiti nel fascicolo "
          f"{sum(1 for x in referto if x.get('avviso_instradato'))}"
          f"   (compaiono fra le comunicazioni del fascicolo)"
          if avvisi_soli else ""),
         (f"     rimasti senza fascicolo   "
          f"{sum(1 for x in referto if 'avviso_instradato' in x and not x['avviso_instradato'])}"
          f"   (sono in archivio_pec.json -> avvisi)"
          if avvisi_soli else ""),
             f"  messi da parte (interni) . {da_parte}" +
             ("   senza allegati, da casella INPS a casella INPS: "
              "sono in archivio_pec.json -> accantonati" if da_parte else ""),
             f"  copie della stessa comunicazione: {copie_viste}" +
             ("   <-- ANOMALO: troppe riconosciute sull'oggetto e il minuto, "
              "senza riferimenti: vedi il dettaglio sotto"
              if presi and sum(1 for x in referto
                               if "minuto" in testo(x.get("copia_regola"))) > presi * 0.2
              else ""),
             f"  unite a un fascicolo gia' presente: {unite}",
             f"  messe in coda a un fascicolo: {accodate}",
             f"  scartate come interne ..... {interne}",
             f"  inoltri fra colleghi ...... {inoltri}",
             f"  non da fascicolo .......... "
             f"{sum(1 for x in referto if x.get('declassata'))}"]),
         f"non corrispondenti ......... {saltati}",
         f"elementi illeggibili ....... {errori}"
           + ("   <-- ANOMALO: piu' del 20%, probabile guasto nel codice"
              if tot and errori > tot * 0.2 else ""),
         ""]
    # Perche' sono illeggibili. Il conteggio da solo e' uno scarto muto: la
    # traccia c'e' gia' in ERRORI, mancava solo stamparla. Sessantotto
    # messaggi che spariscono senza motivo sono sessantotto PEC che potrebbero
    # essere ricorsi.
    if ERRORI:
        per_tipo = {}
        for e in ERRORI:
            k = (e.get("punto") or "?") + " \u2014 " + (e.get("tipo") or "?")
            per_tipo.setdefault(k, []).append(e)
        r += ["PERCHE' SONO ILLEGGIBILI:"]
        for k, l in sorted(per_tipo.items(), key=lambda x: -len(x[1])):
            r.append(f"  {len(l):5}  {k}")
            r.append(f"         {testo(l[0].get('messaggio'))[:88]}")
            for e in l[:3]:
                r.append(f"         [{testo(e.get('cartella'))[:18]}] "
                         f"{testo(e.get('oggetto'))[:56]}")
        r.append("")
    # IL NUMERO DELLE COPIE, DA SOLO, NON SI PUO' GIUDICARE. Centosedici copie
    # su centoquarantacinque sono la normalita' se la stessa comunicazione
    # arriva alla casella dedicata, all'avvocato istruttore e alla direzione;
    # sono un guasto se il riconoscimento sta confondendo PEC distinte. Le due
    # cose si distinguono solo dicendo CON CHE REGOLA e DI CHE COSA ogni copia
    # e' copia: oggetto + giorno + riferimenti identifica una comunicazione,
    # oggetto + minuto e' un ripiego che puo' sbagliare.
    copie = [x for x in referto if x.get("copia_regola")]
    if copie:
        rischiose = [x for x in copie if "minuto" in testo(x.get("copia_regola"))]
        sicure = [x for x in copie if "minuto" not in testo(x.get("copia_regola"))]
        prima = [x for x in copie if not x.get("copia_stesso_lancio")]
        r += ["COPIE DELLA STESSA COMUNICAZIONE: come sono state riconosciute",
              f"  {len(sicure):5d}  sull'oggetto, il giorno e i riferimenti "
              "(numero di ruolo, SISCO, comunicazione): identifica la "
              "comunicazione, non sbaglia",
              f"  {len(rischiose):5d}  sull'oggetto e il minuto, senza nessun "
              "riferimento: e' un ripiego, e due PEC diverse arrivate nello "
              "stesso minuto possono confondersi"
              + ("   <-- da controllare" if rischiose else ""),
              f"  {len(prima):5d}  duplicano una comunicazione gia' in archivio "
              "da prima di questo giro (la casella dedicata, l'avvocato "
              "istruttore e la direzione la ricevono tutte)",
              f"  {len(copie) - len(prima):5d}  duplicano una PEC vista adesso, "
              "in un'altra cartella di Outlook"]
        for x in (rischiose or copie)[:6]:
            r.append(f"    {testo(x.get('data'))[:16]}  [{testo(x.get('cartella'))[:16]}]  "
                     f"{testo(x.get('oggetto'))[:52]}")
            r.append(f"        copia di: {testo(x.get('copia_di'))[:60]}")
        r.append("")
    r += ["PERCHE' SONO STATI SCARTATI (i primi motivi):"]
    for mot, n in sorted(motivi_no.items(), key=lambda x: -x[1])[:12]:
        r.append(f"  {n:5d}  {mot}")

    # Se nessuna PEC ha allegati il difetto e' nella lettura di Outlook, non
    # nei criteri: senza questa misura si cercherebbe nel posto sbagliato.
    # image001.png e simili sono la firma di chi scrive, non documenti:
    # contarli fa risultare "con allegati" mail che non ne hanno nessuno
    def _veri(l):
        return [a for a in (l or [])
                if not re.match(r"^(image|oledata|logo|firma|banner|winmail)"
                                r"[0-9._-]*\.(png|jpe?g|gif|bmp|mso|dat)$",
                                testo(a).strip(), re.I)]
    con_all = sum(1 for x in referto if _veri(x.get("allegati")))
    tot_all = sum(len(_veri(x.get("allegati"))) for x in referto)
    # quante PEC con allegati NON hanno prodotto una cartella: se sono molte,
    # qualcosa le sta scartando prima
    fatte = len(LAVORATE)
    # Non produrre una cartella non vuol dire essere andata perduta: una PEC
    # instradata nel suo fascicolo, o riconosciuta come copia, e' finita dove
    # doveva. L'allarme contava quelle come sospette - 3.073 su 3.614 - e
    # gridava al lupo proprio quando il programma funzionava.
    finite_bene = sum(1 for x in referto
                      if x.get("in_coda") or x.get("unita")
                      or x.get("copia_perche") or x.get("esito") == "copia")
    perse = max(0, con_all - fatte - finite_bene)
    if referto:
        r += ["",
              f"ALLEGATI: {con_all} PEC su {len(referto)} ne hanno, {tot_all} in tutto"]
        if finite_bene:
            r += [f"  di cui {finite_bene} finite in un fascicolo o riconosciute come copia"]
        # In RICOGNIZIONE non si scrive niente per definizione: nessuna PEC
        # finisce in una cartella, e l'allarme suonava sempre, al massimo
        # volume ("154 su 274"). Un allarme che suona sempre non si legge piu'.
        if args.ricognizione and con_all:
            r += ["  (in ricognizione non si scrive niente: che nessuna sia finita",
                  "   in una cartella e' il comportamento voluto, non un difetto)"]
        elif con_all and perse > con_all * 0.3:
            r += [f"  ATTENZIONE: {perse} PEC hanno allegati e non sono finite da nessuna",
                  "  parte: ne' una cartella, ne' un fascicolo, ne' una copia.",
                  "  Se il numero e' alto, qualcosa le sta scartando prima."]
        if not con_all:
            r += ["  NESSUNA PEC HA ALLEGATI: non e' normale.",
                  "  O la casella contiene solo avvisi del Portale, oppure Outlook non",
                  "  li sta esponendo. Controlla aprendone una a mano."]
        else:
            esempi = [x for x in referto if _veri(x.get("allegati"))][:5]
            for x in esempi:
                r.append("    " + ", ".join(_veri(x["allegati"])[:4])[:80])

    if ERRORI:
        per_causa = {}
        for e in ERRORI:
            per_causa.setdefault((e["tipo"], e["messaggio"][:70], e["punto"]), []).append(e)
        r += ["", "=" * 70,
              f"ERRORI: {len(ERRORI)} messaggi non elaborati, {len(per_causa)} cause distinte",
              "=" * 70,
              "Ogni causa e' un difetto solo, anche se colpisce molti messaggi.",
              "La riga indicata dice dove intervenire."]
        for (tipo, msg, punto), casi in sorted(per_causa.items(), key=lambda x: -len(x[1])):
            r += ["", f"  {len(casi):5d}  {tipo}: {msg}",
                  f"         {punto}"]
            for c in casi[:3]:
                r.append(f"         es. [{c['cartella']}] {c['oggetto']}")
        r += [""]
    if referto:
        per_cart = {}
        for x in referto:
            per_cart[x.get("cartella","?")] = per_cart.get(x.get("cartella","?"), 0) + 1
        r += ["", "SELEZIONATI PER CARTELLA:"]
        for c, n in sorted(per_cart.items(), key=lambda x: -x[1]):
            r.append(f"  {n:5d}  {c}")
        per_tipo = {}
        for x in referto:
            per_tipo[x.get("tipo","?")] = per_tipo.get(x.get("tipo","?"), 0) + 1
        r += ["", "TIPO DI COMUNICAZIONE:"]
        for c, n in sorted(per_tipo.items(), key=lambda x: -x[1]):
            r.append(f"  {n:5d}  {c}")
        fasi = {}
        for x in referto:
            if x.get("fase"):
                fasi[x["fase"]] = fasi.get(x["fase"], 0) + 1
        if fasi:
            r += ["", "FASI DELLA CONSULENZA TECNICA:"]
            for c, n in sorted(fasi.items(), key=lambda x: -x[1]):
                r.append(f"  {n:5d}  {c}")
        # LE LETTURE. Il punto della piramide e' poterla misurare: quanti fatti
        # ha prodotto ogni lettore e cosa ha cercato senza trovarlo. Se dopo una
        # modifica un lettore cala, si vede qui e si sa quale.
        if CONTA_LETTORI or CONTA_ATTI:
            r += ["", "=" * 70,
                  "LE LETTURE (chi legge e' separato da chi scrive)",
                  f"  {LETTURE['scritte']} dossier scritti nei fascicoli (letture.json)",
                  "  il registro li legge e decide: l'ingestione non tocca nessun campo"]
            if args.ricognizione:
                r += ["  ATTENZIONE: in ricognizione gli allegati non si aprono. I lettori",
                      "  hanno visto solo oggetto e corpo, quindi sentenze, CTU e termini",
                      "  sono sottostimati: per misurarli serve un'elaborazione."]
            if CONTA_ATTI:
                r += ["", "  atto riconosciuto:"]
                for c, n in sorted(CONTA_ATTI.items(), key=lambda x: -x[1]):
                    r.append(f"  {n:5d}  {c}")
            if CONTA_LETTORI:
                r += ["", "  fatti prodotti, per lettore:"]
                for c, n in sorted(CONTA_LETTORI.items(), key=lambda x: -x[1]):
                    r.append(f"  {n:5d}  {c}")
            if CONTA_RIFIUTI:
                r += ["", "  cercato e NON trovato (un rifiuto scritto non e' un dato mancante):"]
                for c, n in sorted(CONTA_RIFIUTI.items(), key=lambda x: -x[1])[:12]:
                    r.append(f"  {n:5d}  {c}")
        conrg = sum(1 for x in referto if (x.get("rif") or {}).get("rg"))
        consis = sum(1 for x in referto if (x.get("rif") or {}).get("sisco"))
        consen = sum(1 for x in referto if (x.get("rif") or {}).get("sentenza"))
        conspe = sum(1 for x in referto if (x.get("rif") or {}).get("spese_legali"))
        concom = sum(1 for x in referto if (x.get("rif") or {}).get("comunicazione"))
        agganciabili = sum(1 for x in referto if x.get("rif"))
        scartate_int = [x for x in referto if x.get("scartata_tardi")]
        if scartate_int:
            r += ["", f"SCARTATE DOPO L'APERTURA: {len(scartate_int)} comunicazioni interne",
                  "  hanno l'aspetto di una PEC perche' viaggiano in risposta a una vera,",
                  "  ma all'ultimo livello non portano nulla di strutturato.",
                  "  Le prime, per controllare che siano davvero da scartare:"]
            for x in scartate_int[:10]:
                r.append(f"    {x['data'][:10]}  {testo(x['mittente'])[:34]:<34}{testo(x['oggetto'])[:52]}")
        decl = [x for x in referto if x.get("declassata")]
        if decl:
            per_motivo = {}
            for x in decl:
                per_motivo.setdefault(x.get("perche_no") or "non specificato", []).append(x)
            r += ["", "=" * 70,
                  f"NON SONO PEC DA FASCICOLO: {len(decl)}",
                  "=" * 70,
                  "Una PEC merita un fascicolo se ha almeno uno di questi segni:",
                  "  1. l'attestazione del gestore con mittente esterno all'Istituto",
                  "  2. la busta del processo telematico",
                  "  3. il modulo di cancelleria del tribunale",
                  "  4. un allegato che e' un atto (ricorso, relata, decreto, procura)",
                  "  5. un numero di ruolo o un SISCO",
                  "Queste non ne hanno nessuno. Raggruppate per motivo:"]
            for motivo, l in sorted(per_motivo.items(), key=lambda x: -len(x[1])):
                r += ["", f"  {len(l):5d}  {motivo}"]
                for x in l[:4]:
                    r.append(f"           {testo(x.get('data'))[:10]}  "
                             f"{testo(x.get('oggetto'))[:56]}")
            r += ["", "  Se fra queste c'e' un ricorso vero, il criterio e' troppo stretto:",
                  "  segnalalo, perche' significa che manca un segno da riconoscere."]
        arrivate = sum(1 for x in referto if x.get("arrivata_ad_accantonare"))
        motivi = {}
        for x in referto:
            if x.get("giudice_altro") and not x.get("arrivata_ad_accantonare"):
                motivi[x.get("perche_non_spostata") or "motivo non registrato"] = \
                    motivi.get(x.get("perche_non_spostata") or "motivo non registrato", 0) + 1
        r += ["", f"  {arrivate} spostate in ALTRI GIUDICI"]
        for m, n in sorted(motivi.items(), key=lambda x: -x[1]):
            r.append(f"  {n} no: {m}")
        altri = sum(1 for x in referto if x.get("giudice_altro"))
        if altri:
            r += ["", f"DI GIUDICI NON SEGUITI: {altri} PEC",
                  "  le comunicazioni arrivano in una casella comune a tutti i giudici:",
                  "  queste riguardano fascicoli tenuti da altri operatori"]
        # Il giudice trovato solo dentro il documento: e' il caso che sfuggiva
        # a tutti i controlli - oggetto generico, mittente collega, nome del
        # giudice sepolto nel PDF. Dichiararlo serve a controllare che la
        # lettura non stia inventando nomi.
        rec = [x for x in referto if x.get("da_recuperare")]
        if rec:
            r += ["", f"ATTI DA RECUPERARE DAL PORTALE: {len(rec)} PEC",
                  "  Avvisano che sui server dell'Istituto c'e' un atto: non hanno",
                  "  allegati e non possono avere riferimenti, perche' l'informazione",
                  "  sta di la'. Va prelevato da chi ha l'accesso. Non sono scarti:",
                  "  sono lavoro, e nel registro stanno in Novita'."]
            for x in rec[:6]:
                r.append(f"    {testo(x.get('data'))[:10]}  "
                         f"{testo(x.get('mittente'))[:28]:30}{testo(x.get('oggetto'))[:44]}")

        ngd = [x for x in referto if x.get("giudice_dal_documento")]
        if ngd:
            r += ["", f"GIUDICE TROVATO DENTRO IL DOCUMENTO: {len(ngd)} PEC",
                  "  oggetto e mittente non lo dicevano: sta nel PDF o nella firma.",
                  "  Controlla che siano nomi veri, non formule di rito."]
            for x in ngd[:8]:
                r.append(f"    {testo(x.get('giudice_dal_documento'))[:24]:26}"
                         f"{testo(x.get('oggetto'))[:54]}")

        conmod = sum(1 for x in referto if x.get("modulo_tribunale"))
        con_busta2 = sum(1 for x in referto if x["allegati"])
        if conmod:
            r += ["", f"MODULO DI CANCELLERIA letto in {conmod} PEC su {con_busta2} con busta"
                  + (f" ({round(100*conmod/con_busta2)}%)" if con_busta2 else ""),
                  "  ruolo, giudice, parti e provvedimento presi dal testo della PEC,",
                  "  senza dipendere dai PDF (che possono essere scansioni illeggibili).",
                  "  Il modulo sta DENTRO la busta: gli avvisi senza allegati non ce l'hanno."]
        conxml = sum(1 for x in referto if x.get("xml_atto"))
        if conxml:
            r += ["", f"XML DELLA BUSTA TELEMATICA letto in {conxml} PEC",
                  "  usato quando il corpo non dice nulla e i PDF sono scansioni"]
        r += ["", "RIFERIMENTI TROVATI NELL'OGGETTO (agganciano la PEC a un fascicolo):",
              f"  {conrg:5d}  numero di ruolo",
              f"  {consis:5d}  codice SISCO",
              f"  {consen:5d}  numero di sentenza",
              f"  {conspe:5d}  pratica spese legali (n/anno/E)",
              f"  {concom:5d}  numero di comunicazione",
              f"  {agganciabili:5d}  agganciabili senza aprire gli allegati "
              f"({round(100*agganciabili/max(len(referto),1))}%)"]
        senza = sum(1 for x in referto if not x["allegati"])
        r += ["", f"  senza allegati (soli avvisi del Portale): {senza} su {len(referto)}"]
        val = sum(1 for x in referto if x.get("notifica") is True)
        vzt = sum(1 for x in referto if x.get("notifica") is False)
        ign = sum(1 for x in referto if x.get("notifica") is None)
        condc = sum(1 for x in referto if x.get("daticert"))
        dadc = sum(1 for x in referto if x.get("fonte_notifica") == "daticert.xml")
        dacorpo = sum(1 for x in referto if x.get("fonte_notifica") == "corpo")
        con_busta = sum(1 for x in referto if x["allegati"])
        r += ["", "VALIDITA' DELLA NOTIFICA:",
              f"  {val:5d}  notificate alla casella degli atti giudiziari",
              f"  {vzt:5d}  notificate ad altra casella -> difetto di notifica",
              f"  {ign:5d}  casella non individuata",
              f"         di cui {sum(1 for x in referto if x.get('inoltro'))} perche' la PEC e' un"
              " inoltro interno: l'attestazione originale non c'e' piu'"]
        if args.ricognizione:
            r += ["", "  (in ricognizione si guardano solo i destinatari Outlook: elaborando",
                  "   si legge daticert.xml, l'attestazione del gestore, che fa fede)"]
        else:
            r += ["", "  Da dove viene la qualificazione:",
                  f"    {dadc:5d}  da daticert.xml, attestazione del gestore",
                  f"    {dacorpo:5d}  dal corpo del messaggio, dove compare il destinatario",
                  "",
                  f"  attestazioni annidate lette: "
              f"{sum(x.get('n_attestazioni') or 0 for x in referto)}"
              f" (ogni inoltro ne aggiunge una; vale la piu' interna)",
              f"  daticert.xml letto in {condc} PEC su {con_busta} che hanno una busta"
                  + (f" ({round(100*condc/con_busta)}%)" if con_busta else ""),
                  f"  le altre {len(referto)-con_busta} sono avvisi senza allegati: nessuna busta,"
                  " nessuna attestazione"]

        # quali caselle ricevono davvero gli atti
        per_casella = {}
        for x in referto:
            c = x.get("casella_notifica")
            if c:
                per_casella[c] = per_casella.get(c, 0) + 1
        if per_casella:
            r += ["", "  CASELLE DI DESTINAZIONE ACCERTATE:"]
            for c, n in sorted(per_casella.items(), key=lambda y: -y[1]):
                deputata = any(v.lower() == c.lower() for v in CASELLE.get("valida", []))
                r.append(f"    {n:5d}  {c}" + ("   <- casella deputata" if deputata else ""))
        per_mot = {}
        for x in referto:
            per_mot[x["motivo"]] = per_mot.get(x["motivo"], 0) + 1
        r += ["", "CRITERIO CHE LI HA FATTI SCATTARE:"]
        for c, n in sorted(per_mot.items(), key=lambda x: -x[1])[:12]:
            r.append(f"  {n:5d}  {c}")
        firme = {}
        for x in referto:
            f = (x["oggetto"].lower().strip(), x["data"][:10])
            firme.setdefault(f, []).append(x)
        ripetuti = {k: v for k, v in firme.items() if len(v) > 1}
        if ripetuti:
            r += ["", "PROBABILI DOPPIONI (stesso oggetto, stesso giorno):",
                  f"  {sum(len(v)-1 for v in ripetuti.values())} messaggi in eccesso "
                  f"su {len(ripetuti)} comunicazioni distinte",
                  "  (la stessa PEC arriva alla casella dedicata e a quella dell'avvocato istruttore)"]
            for (ogg, gg), v in sorted(ripetuti.items(), key=lambda x: -len(x[1]))[:8]:
                r.append(f"  {len(v)}x  {gg}  {ogg[:70]}")
                r.append(f"        cartelle: {', '.join(sorted({y['cartella'] for y in v}))}")
        if sfiorati:
            r += ["", "=" * 70,
                  "SCARTATI PER UN SOFFIO (un solo indizio): controlla se ne perdi di buoni"]
            for x in sfiorati[:25]:
                r += ["", f"  {x['data']}  [{x['cartella']}]  {x['motivo']}",
                      f"  oggetto : {x['oggetto']}"]
        r += ["", "=" * 70, "MESSAGGI SELEZIONATI (primi 60):"]
        for x in referto[:60]:
            r += ["", f"  {x['data']}  {x['mittente']}   [{x.get('tipo','?')}]",
                  f"  oggetto : {x['oggetto']}",
                  f"  motivo  : {x['motivo']}",
                  f"  allegati: {', '.join(x['allegati']) if x['allegati'] else '(nessuno)'}"]
    testo_ref = "\n".join(r)
    print("\n" + testo_ref)
    out = BASE / "referto_ingestione.txt"
    out.write_text(testo_ref, encoding="utf-8")
    log(f"referto salvato in {out}")
    diario("OK", f"{tot} letti, {presi} ai criteri, {len(referto)} lavorati, "
                 f"{len(LAVORATE)} cartelle, {len(IN_CODA)} in coda")
    return 0

def servizio(minuti):
    """Resta in ascolto e rilegge la posta ogni tot minuti.

    Il registro ricontrolla da solo la cartella IN ARRIVO ogni cinque minuti,
    ma quella cartella la riempie questo programma: se nessuno lo lancia, la
    rilettura continua a trovare la posta di stamattina e le PEC arrivate
    dopo restano in Outlook. Da qui in avanti il giro e' completo - uno tiene
    la casella, l'altro tiene la vista.

    Un lancio che fallisce NON ferma il servizio: OneDrive che si riallinea,
    Outlook che si riavvia, la rete che manca un minuto sono normali in un
    ufficio. Si scrive nel diario e si riprova al giro dopo.
    """
    import time
    giro = 0
    # Nessuna riga nel diario all'avvio: il diario e' l'elenco delle LETTURE
    # della posta, e il registro giudica lo stato dall'ultima riga. Una riga
    # "SERVIZIO" in cima faceva sparire il riquadro calmo appena costruito e
    # mostrava l'avvio riuscito col rosso dei giri falliti. A parlare sia il
    # primo giro, che e' la prima lettura vera.
    log(f"servizio avviato: rilettura della posta ogni {minuti} minuti. "
        f"Ctrl+C per fermarlo.")
    while True:
        giro += 1
        inizio = datetime.now()
        log("")
        log("=" * 70)
        log(f"giro {giro} \u2014 {inizio:%d/%m/%Y %H:%M:%S}")
        log("=" * 70)
        try:
            main()
        except KeyboardInterrupt:
            raise
        except Exception as e:
            import traceback
            tb = traceback.format_exc().strip().split("\n")
            diario("FALLITO", f"giro {giro}: {type(e).__name__}: {e}")
            err(f"giro {giro} fallito: {e}")
            for r in tb[-3:]:
                err("    " + r.strip())
            err("il servizio continua: riprovo al prossimo giro")
        # l'attesa si conta dalla FINE del giro: un lancio che dura tre minuti
        # non deve accorciare la pausa a due
        prossimo = datetime.now() + timedelta(minutes=minuti)
        log(f"prossimo giro alle {prossimo:%H:%M}")
        try:
            time.sleep(minuti * 60)
        except KeyboardInterrupt:
            raise


if __name__ == "__main__":
    # Il diario deve registrare ANCHE i fallimenti: un lancio automatico che
    # esplode a meta' non scrive referto e non lascia traccia, e passano
    # settimane prima che qualcuno si accorga che la posta non entra piu'.
    try:
        # --ogni trasforma il lancio in un servizio: si intercetta prima di
        # main(), che resta il singolo giro e non sa di essere ripetuto.
        _ogni = None
        _via = []          # le POSIZIONI da togliere, non i valori: con
                           # "--ogni 5 --altro 5" togliere ogni "5" portava
                           # via anche quello dell'altro argomento
        for _i, _a in enumerate(sys.argv):
            if _a == "--ogni" and _i + 1 < len(sys.argv):
                _ogni = sys.argv[_i + 1]
                _via += [_i, _i + 1]
            elif _a.startswith("--ogni="):
                _ogni = _a.split("=", 1)[1]
                _via.append(_i)
        if _ogni:
            try:
                _m = int(_ogni)
            except ValueError:
                err(f"--ogni vuole un numero di minuti, non \"{_ogni}\"")
                sys.exit(2)
            if _m < 1:
                err("--ogni vuole almeno un minuto")
                sys.exit(2)
            sys.argv = [_x for _j, _x in enumerate(sys.argv) if _j not in _via]
            servizio(_m)
            sys.exit(0)
        _uscita = main()
    except KeyboardInterrupt:
        diario("INTERROTTO", "fermato da chi lo stava guardando")
        raise
    except Exception as _e:
        import traceback
        _tb = traceback.format_exc().strip().split("\n")
        diario("FALLITO", f"{type(_e).__name__}: {_e} | {_tb[-2].strip() if len(_tb) > 1 else ''}")
        err(f"lancio fallito: {_e}")
        for _r in _tb[-3:]:
            err("    " + _r.strip())
        sys.exit(2)
    if _uscita not in (0, None):
        diario("USCITA " + str(_uscita),
               ULTIMO_ERRORE or "il programma si e' fermato da solo")
    sys.exit(_uscita)
