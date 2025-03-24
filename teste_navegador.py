from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import time
import os
import pyautogui
import keyboard
import json
from datetime import datetime

def carregar_credenciais(nome_arquivo):
    """Carrega as credenciais de e-mail do arquivo especificado."""
    try:
        # Obtém o diretório do script atual
        script_dir = os.path.dirname(os.path.abspath(__file__))
        nome_arquivo = os.path.join(script_dir, nome_arquivo)

        # Verifica se o arquivo existe
        if not os.path.exists(nome_arquivo):
            raise FileNotFoundError(f"O arquivo {nome_arquivo} não foi encontrado no diretório do script.")

        # Lê o arquivo e carrega as credenciais
        with open(nome_arquivo, "r", encoding="utf-8") as arquivo:
            linhas = arquivo.readlines()
            if len(linhas) < 4:
                raise ValueError("O arquivo deve conter pelo menos 4 linhas com as credenciais.")
            
            email_imap = linhas[0].strip()  # Primeira linha: e-mail
            senha_app = linhas[1].strip()   # Segunda linha: senha app
            usuario_portal = linhas[2].strip()  # Terceira linha: usuário portal
            senha_portal = linhas[3].strip()    # Quarta linha: senha portal

            print("Credenciais carregadas com sucesso.")
            return email_imap, senha_app, usuario_portal, senha_portal

    except Exception as e:
        print(f"Erro ao carregar credenciais do arquivo {nome_arquivo}: {e}")
        return None, None, None, None

def esperar_elemento(navegador, xpath, timeout=10):
    try:
        return WebDriverWait(navegador, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))
    except TimeoutException:
        return None

def gravar_cliques():
    """Função para gravar cliques do mouse e salvá-los em um arquivo."""
    cliques = []
    contador = 1  # Contador para numeração crescente
    print("\nGravador de cliques iniciado!")
    print("Pressione 'ESC' para parar a gravação")
    print("Pressione 'Scroll Lock' para registrar um clique")
    
    def on_scroll_lock():
        nonlocal contador
        x, y = pyautogui.position()
        timestamp = datetime.now().strftime("%H:%M:%S")
        cliques.append({
            "numero": contador,
            "x": x,
            "y": y,
            "hora": timestamp
        })
        print(f"Clique {contador} registrado: ({x}, {y}) às {timestamp}")
        contador += 1
    
    def on_esc():
        return True
    
    # Registra os handlers dos eventos
    keyboard.on_press_key("scroll lock", lambda _: on_scroll_lock())
    keyboard.wait("esc")
    
    # Salva os cliques em um arquivo JSON
    if cliques:
        nome_arquivo = f"cliques_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(nome_arquivo, "w", encoding="utf-8") as f:
            json.dump(cliques, f, indent=4, ensure_ascii=False)
        print(f"\nCliques salvos em: {nome_arquivo}")
        print(f"Total de cliques registrados: {contador-1}")
    else:
        print("\nNenhum clique foi registrado.")

# Configurar opções do Firefox
firefox_options = Options()
firefox_options.add_argument("--start-maximized")
firefox_options.add_argument("--disable-notifications")
firefox_options.add_argument("--disable-popup-blocking")
firefox_options.add_argument("--disable-gpu")
firefox_options.add_argument("--no-sandbox")
firefox_options.add_argument("--disable-dev-shm-usage")

# Especificar o caminho do Firefox manualmente
firefox_options.binary_location = r"C:\Program Files\Mozilla Firefox\firefox.exe"

try:
    print("Iniciando configuração do GeckoDriver...")
    # Instala o GeckoDriver
    service = Service(GeckoDriverManager().install())
    print(f"GeckoDriver instalado em: {service.path}")
    
    print("Iniciando o navegador Firefox...")
    # Inicializa o navegador com as opções configuradas
    navegador = webdriver.Firefox(service=service, options=firefox_options)
    print("Navegador Firefox iniciado com sucesso!")
    
    print("Maximizando a janela...")
    navegador.maximize_window()
    
    print("Acessando o site...")
    navegador.get("https://www.portaldeboletos.com.br/brmarinas")
    print("Site acessado com sucesso!")

    # Carrega as credenciais
    email_imap, senha_app, usuario_portal, senha_portal = carregar_credenciais("login_imap.txt")
    
    if not all([email_imap, senha_app, usuario_portal, senha_portal]):
        raise ValueError("Não foi possível carregar as credenciais.")

    # Preenche o campo de usuário
    elemento_usuario = esperar_elemento(navegador, '//*[@id="frm_tab_usuario_senha"]/div/input[2]')
    if elemento_usuario:
        elemento_usuario.send_keys(usuario_portal)
        print("Usuário preenchido com sucesso!")
    else:
        print("Campo de usuário não encontrado!")

    # Preenche o campo de senha
    elemento_senha = esperar_elemento(navegador, '//*[@id="frm_tab_usuario_senha"]/div/input[3]')
    if elemento_senha:
        elemento_senha.send_keys(senha_portal)
        print("Senha preenchida com sucesso!")
    else:
        print("Campo de senha não encontrado!")

    # Clica no botão de acessar
    elemento_acessar = esperar_elemento(navegador, '//*[@id="acessar1"]')
    if elemento_acessar:
        elemento_acessar.click()
        print("Botão de acessar clicado com sucesso!")
    else:
        print("Botão de acessar não encontrado!")

    # Inicia o gravador de cliques
    print("\nIniciando gravador de cliques...")
    gravar_cliques()

    # Aguarda 30 minutos antes de fechar
    print("Aguardando 30 minutos...")
    time.sleep(1800)  # 30 minutos = 1800 segundos
    
except Exception as e:
    print(f"Erro ao executar o script: {e}")
    print("Detalhes do erro:", str(e))
    raise


 
finally:
    # Fecha o navegador
    print("Fechando o navegador...")
    navegador.quit()
    print("Navegador fechado com sucesso!") 