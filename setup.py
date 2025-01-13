import os
import subprocess
import sys

def create_virtualenv():
    # Nome do ambiente virtual
    venv_name = "Instalador"

    # Verifica se o ambiente já existe
    if not os.path.exists(venv_name):
        print("Criando o ambiente virtual...")
        subprocess.check_call([sys.executable, "-m", "venv", venv_name])
        print("Ambiente virtual criado com sucesso!")
    else:
        print("Ambiente virtual já existe.")

def install_dependencies():
    # Caminho para o pip dentro do ambiente virtual
    pip_path = os.path.join("venv", "Scripts", "pip") if os.name == "nt" else os.path.join("venv", "bin", "pip")

    # Instala as dependências do arquivo requirements.txt
    if os.path.exists("requirements.txt"):
        print("Instalando dependências...")
        subprocess.check_call([pip_path, "install", "-r", "requirements.txt"])
        print("Dependências instaladas com sucesso!")
    else:
        print("Arquivo requirements.txt não encontrado.")

if __name__ == "__main__":
    create_virtualenv()
    install_dependencies()
