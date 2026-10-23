import sys
import runpy

print("--- AVVIO PEC INTAKE IN MODALITA' CLOUD (Emulatore IMAP) ---")
try:
    runpy.run_path('pec_intake.py', run_name='__main__')
except SystemExit as e:
    sys.exit(e.code)
except Exception as e:
    print(f"Errore critico imprevisto: {e}")
    sys.exit(1)
