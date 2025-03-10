import subprocess
import sys

try:
    # Executa relacionamentos.py usando o mesmo interpretador Python
    subprocess.run([sys.executable, "relacionamentos.py"], check=True)
    
    # Executa ver1.1.py após conclusão bem-sucedida do primeiro
    subprocess.run([sys.executable, "ver1.1.py"], check=True)

except subprocess.CalledProcessError as e:
    print(f"Erro na execução do script: {e}", file=sys.stderr)
    sys.exit(1)
except Exception as e:
    print(f"Erro inesperado: {str(e)}", file=sys.stderr)
    sys.exit(1)