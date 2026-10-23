import sys
import os
import runpy

# 1. Forziamo Python a credere che win32com (Outlook) NON esista.
sys.modules['win32com'] = None
sys.modules['win32com.client'] = None
sys.modules['pythoncom'] = None

# 2. Impostiamo l'ambiente
os.environ['GITHUB_ACTIONS'] = 'true'

print("--- AVVIO PEC INTAKE IN MODALITA' CLOUD ---")
try:
    # run_path esegue il file come se lo avessi lanciato direttamente da terminale.
    # Passa automaticamente gli argomenti della riga di comando (es. --dal 2026-09-01).
    runpy.run_path('pec_intake.py', run_name='__main__')
except SystemExit as e:
    sys.exit(e.code)
except Exception as e:
    print(f"Errore critico imprevisto: {e}")
    sys.exit(1)
