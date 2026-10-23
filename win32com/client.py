import imaplib
import email
import os
from email.header import decode_header
import re
from datetime import datetime
from email.utils import parsedate_to_datetime

class Dispatch:
    def __init__(self, prog_id):
        if prog_id == "Outlook.Application":
            self.app = OutlookApp()
        else:
            self.app = None
    def __getattr__(self, name):
        if self.app: return getattr(self.app, name)
        return MagicDummy()

class MagicDummy:
    def __getattr__(self, name): return MagicDummy()
    def __call__(self, *args, **kwargs): return MagicDummy()

class OutlookApp:
    def GetNamespace(self, name):
        return Namespace()

class Namespace:
    def __init__(self):
        server = os.environ.get("PEC_IMAP_SERVER", "imaps.legalmail.it")
        user = os.environ.get("PEC_USER", "")
        pwd = os.environ.get("PEC_PASSWORD", "")
        
        try:
            self.imap = imaplib.IMAP4_SSL(server)
            self.imap.login(user, pwd)
        except Exception as e:
            print(f"ERRORE CONNESSIONE IMAP: {e}")
            raise e
            
        self.Folders = FolderCollection(self.imap)

    def GetDefaultFolder(self, num):
        return Folder("INBOX", self.imap)

    def Logon(self, *args, **kwargs): pass

class FolderCollection:
    def __init__(self, imap):
        self.imap = imap
        typ, data = self.imap.list()
        self.folders = []
        if typ == 'OK':
            for line in data:
                # Estrae il nome della cartella, gestendo le virgolette
                match = re.search(r'"([^"]*)"$', line.decode('utf-8'))
                if match:
                    self.folders.append(match.group(1))
        self.Count = len(self.folders)
    
    def Item(self, index):
        return Folder(self.folders[index-1], self.imap)
    
    def __iter__(self):
        for f in self.folders:
            yield Folder(f, self.imap)

class Folder:
    def __init__(self, name, imap):
        self.Name = name
        self.FolderPath = name
        self.imap = imap
        self.Items = Items(name, imap)

class Items:
    def __init__(self, folder_name, imap):
        self.folder_name = folder_name
        self.imap = imap
        self.message_ids = []
        self.current_idx = 0
        self._fetch_all()

    def _fetch_all(self):
        try:
            self.imap.select(f'"{self.folder_name}"', readonly=True)
            typ, data = self.imap.search(None, 'ALL')
            if typ == 'OK':
                self.message_ids = data[0].split()
        except Exception:
            pass

    def Restrict(self, filter_str):
        match = re.search(r"'(\d{1,2})/(\d{1,2})/(\d{4})'", filter_str)
        if match:
            d, m, y = match.groups()
            # Formatta gg-mmm-aaaa per IMAP
            mesi = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            data_imap = f"{int(d):02d}-{mesi[int(m)-1]}-{y}"
            try:
                self.imap.select(f'"{self.folder_name}"', readonly=True)
                typ, data = self.imap.search(None, f'SINCE {data_imap}')
                if typ == 'OK':
                    self.message_ids = data[0].split()
            except Exception:
                pass
        self.current_idx = 0
        return self

    def Sort(self, *args, **kwargs):
        pass # IMAP restituisce già ordinato per data

    def GetFirst(self):
        self.current_idx = 0
        return self._get_current()

    def GetNext(self):
        self.current_idx += 1
        return self._get_current()

    def _get_current(self):
        if self.current_idx >= len(self.message_ids): return None
        msg_id = self.message_ids[self.current_idx]
        try:
            typ, data = self.imap.fetch(msg_id, '(RFC822)')
            if typ == 'OK' and data[0] is not None:
                return Message(data[0][1])
        except Exception as e:
            print(f"Errore fetch: {e}")
            pass
        return None

    def __len__(self):
        return len(self.message_ids)

class Message:
    def __init__(self, raw):
        self.msg = email.message_from_bytes(raw)
        self.Subject = self._decode(self.msg.get('Subject', ''))
        self.SenderName = self._decode(self.msg.get('From', ''))
        self.To = self._decode(self.msg.get('To', ''))
        
        # Gestione della data
        date_str = self.msg.get('Date', '')
        try:
            self.ReceivedTime = parsedate_to_datetime(date_str)
        except:
            self.ReceivedTime = datetime.now()
            
        self.Body = ""
        self.Attachments = AttachmentCollection()
        self._parse()

    def _decode(self, header):
        if not header: return ""
        decoded = decode_header(header)
        return "".join([t[0].decode(t[1] or 'utf-8', errors='ignore') if isinstance(t[0], bytes) else t[0] for t in decoded])

    def _parse(self):
        for part in self.msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                payload = part.get_payload(decode=True)
                if payload:
                    self.Body += payload.decode(part.get_content_charset() or 'utf-8', errors='ignore') + "\n"
            else:
                filename = part.get_filename()
                if filename:
                    self.Attachments._add(Attachment(filename, part.get_payload(decode=True)))

class AttachmentCollection:
    def __init__(self):
        self.attachments = []
        self.Count = 0
    def _add(self, att):
        self.attachments.append(att)
        self.Count += 1
    def Item(self, index):
        return self.attachments[index-1]

class Attachment:
    def __init__(self, filename, payload):
        self.FileName = filename
        self._payload = payload
    def SaveAsFile(self, path):
        with open(path, 'wb') as f:
            f.write(self._payload)
