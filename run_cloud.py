import sys
import builtins
import runpy

# Importiamo il nostro finto Outlook
import win32com.client

# Lo iniettiamo nel namespace globale di Python, 
# così pec_intake.py lo trova anche se ha l'import commentato!
builtins.win32com = win32com

print("--- AVVIO PEC INTAKE IN MODALITA' CLOUD (Emulatore IMAP) ---")
try:
    # Eseguiamo lo script originale
    runpy.run_path('pec_intake.py', run_name='__main__')
except SystemExit as e:
    sys.exit(e.code)
except Exception as e:
    print(f"Errore critico imprevisto: {e}")
    sys.exit(1)
