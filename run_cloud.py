import sys
import os
import subprocess

# Creiamo un file temporaneo che dice a Python di non usare Outlook
with open("blocca_outlook.py", "w") as f:
    f.write("import sys\nsys.modules['win32com'] = None\nsys.modules['win32com.client'] = None\n")

# Prepariamo il comando per lanciare il tuo script originale
# Passiamo gli argomenti della riga di comando a pec_intake.py
args = sys.argv[1:] if len(sys.argv) > 1 else ["--dal", "2026-09-01"]
cmd = [sys.executable, "-c", "exec(open('pec_intake.py').read())"] + args

# Prepariamo l'ambiente finto
env = os.environ.copy()
env['PYTHONSTARTUP'] = 'blocca_outlook.py'

print("--- AVVIO PEC INTAKE IN MODALITA' CLOUD ---")
result = subprocess.run(cmd, env=env)
sys.exit(result.returncode)
