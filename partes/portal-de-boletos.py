from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.firefox.options import Options
from email.header import decode_header
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

import keyboard
import re
import pyautogui
import sys
import time
import email
import imaplib
import os
import re
import sys
import locale
import pyautogui

delay = 0.7
tempo = 1
fast = 0.2
wait = 2

##################################################################################################################################################
# Função para monitorar interrupções
def monitorar_interrupcao():
    print("Monitorando interrupções... Pressione 'ESC' para encerrar.")
    keyboard.wait('esc')  # Aguarda o pressionamento da tecla 'ESC'
    print("Interrupção detectada! Encerrando o script...")
    encerrar_codigo()
# Função para encerrar o código por completo
def encerrar_codigo():
    try:
        if 'navegador' in globals() and navegador:
            navegador.quit()
            encerrar_codigo()  # Fecha o navegador do Selenium
        print("Navegador fechado.")
    except Exception as e:
        print(f"Erro ao fechar o navegador: {e}")
    
    # Finaliza o script completamente
    os._exit(0)
##################################################################################################################################################

def encontrar_arquivo(nome_arquivo, diretorio_base):
    """
    Procura por um arquivo específico em um diretório e seus subdiretórios.
    
    :param nome_arquivo: Nome do arquivo a ser encontrado.
    :param diretorio_base: Diretório base onde a busca será iniciada.
    :return: Caminho completo do arquivo, se encontrado. Caso contrário, None.
    """
    for raiz, _, arquivos in os.walk(diretorio_base):
        if nome_arquivo in arquivos:
            return os.path.join(raiz, nome_arquivo)
    return None
# Obtém o diretório do script atual
script_dir = os.path.dirname(os.path.abspath(__file__))
nome_arquivo = os.path.join(script_dir, "login_imap.txt")

# Busca o arquivo no diretório do script e subdiretórios
nome_arquivo = encontrar_arquivo(nome_arquivo, script_dir)

def limpar_pasta(caminho_pasta):
    """
    Remove todos os arquivos da pasta especificada.
    :param caminho_pasta: str - Caminho da pasta a ser limpa
    """
    try:
        for arquivo in os.listdir(caminho_pasta):
            nome_arquivo = os.path.join(caminho_pasta, arquivo)
            if os.path.isfile(nome_arquivo):
                os.remove(nome_arquivo)
                print(f"Arquivo removido: {nome_arquivo}")
        print("Pasta limpa com sucesso.")
    except Exception as e:
        print(f"Erro ao limpar a pasta: {e}")


def obter_caminho_pasta(nome_arquivo):
    """Obtém o caminho da pasta especificada na linha 5 do arquivo."""
    try:
        # Obtém o diretório onde o script está localizado
        script_dir = os.path.dirname(os.path.abspath(__file__))
        nome_arquivo = os.path.join(script_dir, nome_arquivo)

        # Abre o arquivo no diretório do script
        with open(nome_arquivo, 'r', encoding='utf-8') as arquivo:
            linhas = arquivo.readlines()
            if len(linhas) >= 5:
                return linhas[4].strip()
            else:
                print("Elemento 'Usuários Pagadores' não encontrado!")
                raise ValueError("O arquivo não contém 5 linhas.")
    except Exception as e:
        print(f"Erro ao ler o arquivo {nome_arquivo}: {e}")
        return None

# Uso da função
caminho_pasta = obter_caminho_pasta("login_imap.txt")

# Verifica se o caminho foi obtido com sucesso e limpa a pasta
if caminho_pasta:
    limpar_pasta(caminho_pasta)
else:
    print("Erro ao obter o caminho da pasta. Verifique o arquivo de configuração.")

# Função para decodificar o nome do arquivo
def decodificar_nome_arquivo(nome_arquivo):
    """Decodifica o nome do arquivo, caso ele esteja codificado."""
    if nome_arquivo:
        partes = decode_header(nome_arquivo)
        nome_decodificado = ''.join(
            parte.decode(encoding or 'utf-8') if isinstance(parte, bytes) else parte 
            for parte, encoding in partes
        )
        return nome_decodificado
    return None

# Função para buscar e baixar um anexo específico por e-mail
def baixar_anexo_por_email(usuario_portal, senha_portal, nome_arquivo_procurado, pasta_destino):
    """Busca e baixa um anexo específico de um e-mail."""
    try:
        # Conectar ao servidor IMAP
        servidor = imaplib.IMAP4_SSL("imap.gmail.com")
        servidor.login(usuario_portal, senha_portal)
        servidor.select("inbox")  # Seleciona a caixa de entrada

        # Buscar mensagens na caixa de entrada
        status, mensagens = servidor.search(None, "ALL")
        if status != "OK":
            print("Erro ao buscar mensagens.")
            return False

        ids_mensagens = mensagens[0].split()

        # Iterar sobre os IDs das mensagens (do mais recente ao mais antigo)
        for id_mensagem in reversed(ids_mensagens):
            status, msg_data = servidor.fetch(id_mensagem, "(RFC822)")
            if status != "OK":
                continue

            for resposta in msg_data:
                if isinstance(resposta, tuple):
                    msg = email.message_from_bytes(resposta[1])
                    # Verificar anexos na mensagem
                    if msg.is_multipart():
                        for parte in msg.walk():
                            if parte.get_content_disposition() == "attachment":
                                nome_arquivo = decodificar_nome_arquivo(parte.get_filename())
                                if nome_arquivo == nome_arquivo_procurado:
                                    caminho_completo = os.path.join(pasta_destino, nome_arquivo)
                                    with open(caminho_completo, "wb") as f:
                                        f.write(parte.get_payload(decode=True))
                                    print(f"Anexo salvo: {caminho_completo}")
                                    servidor.logout()
                                    return True

        print("Nenhum anexo encontrado com o nome especificado.")
        servidor.logout()
        return False

    except Exception as e:
        print(f"Erro ao acessar o e-mail: {e}")
        return False

# Função para carregar credenciais do arquivo login_imap.txt

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
            if len(linhas) < 2:
                raise ValueError("O arquivo deve conter pelo menos 2 linhas: e-mail na primeira linha e senha na segunda.")
            
            usuario_portal = linhas[0].strip()  # Primeira linha: e-mail
            senha_portal = linhas[1].strip()   # Segunda linha: senha

            print("Credenciais carregadas com sucesso.")
            return usuario_portal, senha_portal

    except Exception as e:
        print(f"Erro ao carregar credenciais do arquivo {nome_arquivo}: {e}")
        return None, None


# Configurações do script
nome_arquivo = "login_imap.txt"
usuario_portal, senha_portal = carregar_credenciais(nome_arquivo)

if not usuario_portal or not senha_portal:
    print("Erro ao carregar credenciais. Encerrando...")
else:
    nome_arquivo_procurado = "pagadores.csv"
    pasta_destino = os.path.expanduser("~/Downloads")  # Altere conforme necessário

    # Executar a função para baixar o anexo
    print("Buscando e baixando o anexo...")
    sucesso = baixar_anexo_por_email(usuario_portal, senha_portal, nome_arquivo_procurado, pasta_destino)
    if sucesso:
        print("Anexo baixado com sucesso!")
    else:
        print("Falha ao baixar o anexo.")

    # Função para extrair um código numérico de 6 dígitos do texto fornecido
    def extrair_codigo(texto):
        """Extrai um código numérico de 6 dígitos do texto fornecido."""
        padrao = r"\b\d{6}\b"
        match = re.search(padrao, texto)
        return match.group(0) if match else None

# Função para conectar ao servidor de e-mail
def conectar_email(usuario_portal, senha_portal):
    """Conecta ao servidor de e-mail IMAP e retorna a conexão."""
    try:
        servidor = imaplib.IMAP4_SSL("imap.gmail.com")
        servidor.login(usuario_portal, senha_portal)
        return servidor
    except Exception as e:
        print(f"Erro ao conectar ao servidor de e-mail: {e}")
        return None

# Função para buscar o código de confirmação no e-mail
def buscar_codigo_6(servidor, remetente_filtro):
    """Busca o código de confirmação no e-mail do remetente especificado."""
    try:
        servidor.select("inbox")
        status, mensagens = servidor.search(None, f'(FROM "{remetente_filtro}")')
        if status != "OK":
            print("Erro ao buscar mensagens.")
            return None

        ids_mensagens = mensagens[0].split()
        for id_mensagem in reversed(ids_mensagens):
            status, msg_data = servidor.fetch(id_mensagem, "(RFC822)")
            if status != "OK":
                continue

            for resposta in msg_data:
                if isinstance(resposta, tuple):
                    msg = email.message_from_bytes(resposta[1])
                    if msg.is_multipart():
                        for parte in msg.walk():
                            if parte.get_content_type() == "text/plain":
                                conteudo = parte.get_payload(decode=True).decode()
                                codigo = extrair_codigo(conteudo)
                                if codigo:
                                    return codigo
        return None
    except Exception as e:
        print(f"Erro ao buscar o código de confirmação: {e}")
        return None

def preencher_campo_xpath(navegador, xpath_campo, valor):
    """Preenche um campo no navegador identificado pelo XPath."""
    try:
        campo = WebDriverWait(navegador, 5).until(  # Increased wait time from 5 to 10 seconds
        )
        campo.send_keys(valor)
    except Exception as e:
        print(f"Erro ao preencher o campo: {e}")

# Configurações de ambiente
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '2'
sys.stdout.reconfigure(encoding='utf-8')  # Configura saída padrão para UTF-8

try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')  # Configura o locale para português do Brasil
except locale.Error as e:
    print(f"Erro ao configurar o locale: {e}")

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
    
except Exception as e:
    print(f"Erro ao iniciar o Firefox: {e}")
    print("Detalhes do erro:", str(e))
    raise

script_dir = os.path.dirname(os.path.abspath(__file__))
nome_arquivo = os.path.join(script_dir, "login_imap.txt")

# Verifica se o arquivo existe
if not os.path.exists(nome_arquivo):
    raise FileNotFoundError(f"Arquivo {nome_arquivo} não encontrado.")

# Lê as credenciais do arquivo
with open(nome_arquivo, "r", encoding='utf-8') as arquivo:
    linhas = arquivo.readlines()
    email_imap = linhas[0].strip()
    senha_app = linhas[1].strip()
    usuario_portal = linhas[2].strip()
    senha_portal = linhas[3].strip()

print("Credenciais carregadas com sucesso!")

def esperar_elemento(navegador, xpath, timeout=10):
    try:
        return WebDriverWait(navegador, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))
    except TimeoutException:
        return None

# (Usuário)
elemento_usuario_portal = esperar_elemento(navegador, '//*[@id="frm_tab_usuario_senha"]/div/input[2]')
if elemento_usuario_portal:
    elemento_usuario_portal.send_keys(usuario_portal)
else:
    print("Campo 'Usuário' não encontrado!")

# (Senha)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="frm_tab_usuario_senha"]/div/input[3]')
if elemento_senha_portal:
    elemento_senha_portal.send_keys(senha_portal)
else:
    print("Campo 'Senha' não encontrado!")

# (Botão Acessar) - primeiro
elemento_acessar = esperar_elemento(navegador, '//*[@id="acessar1"]')
if elemento_acessar:
    elemento_acessar.click()
else:
    print("Botão 'Acessar' não encontrado!")

#   função para Verificar se tem alguma notificação de alerta clicar e continuar caso nn tenha continuar normalmente
def verificar_e_clicar(navegador, xpath):
    try:
        # Aguarda até que o elemento esteja presente no DOM (tempo máximo de 5 segundos)
        elemento = WebDriverWait(navegador, 5).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        # Clica no elemento caso ele seja encontrado
        elemento.click()
        print("Elemento encontrado e clicado.")
    except:
        # Caso o elemento não seja encontrado, continua sem erros
        print("Elemento não encontrado. Continuando a automação.")

# Verifica se tem alguma notificação de satisfação
verificar_e_clicar(navegador, '//*[@id="indecx-widget-close-btn"]')

# Até aqui ta Ok

time.sleep(1)


def conectar_email(usuario_portal, senha_portal):
    """Conecta ao servidor de e-mail IMAP e retorna a conexão."""
    try:
        servidor = imaplib.IMAP4_SSL("imap.gmail.com")
        servidor.login(usuario_portal, senha_portal)
        return servidor
    except Exception as e:
        print(f"Erro ao conectar ao servidor de e-mail: {e}")
        return None

# Função para decodificar o assunto do e-mail
def decodificar_assunto(assunto):
    """Decodifica o assunto do e-mail."""
    partes = decode_header(assunto)
    assunto_decodificado = ''.join(
        parte.decode(encoding or 'utf-8') if isinstance(parte, bytes) else parte
        for parte, encoding in partes
    )
    return assunto_decodificado

# Função para extrair um código numérico de 6 dígitos do texto fornecido
def extrair_codigo(texto):
    """Extrai um código numérico de 6 dígitos do texto fornecido."""
    padrao = r"\b\d{6}\b"
    match = re.search(padrao, texto)
    return match.group(0) if match else None

# Função para buscar o código no e-mail com assunto específico
def buscar_codigo_por_assunto(servidor, assunto_filtro):
    """Busca o código de confirmação em um e-mail com o assunto especificado."""
    try:
        # Adiciona um atraso de 10 segundos antes de iniciar a busca
        time.sleep(5)

        servidor.select("inbox")
        status, mensagens = servidor.search(None, "ALL")
        if status != "OK" or not mensagens[0]:
            print("Nenhum e-mail encontrado.")
            return None

        ids_mensagens = mensagens[0].split()

        # Iterar do mais recente para o mais antigo
        for id_mensagem in reversed(ids_mensagens):
            status, msg_data = servidor.fetch(id_mensagem, "(RFC822)")
            if status != "OK":
                continue

            for resposta in msg_data:
                if isinstance(resposta, tuple):
                    msg = email.message_from_bytes(resposta[1])

                    # Decodificar o assunto do e-mail
                    assunto = msg["Subject"]
                    if assunto:
                        assunto_decodificado = decodificar_assunto(assunto)
                        if assunto_filtro in assunto_decodificado:
                            # Processar o corpo do e-mail para encontrar o código
                            if msg.is_multipart():
                                for parte in msg.walk():
                                    if parte.get_content_type() == "text/plain":
                                        conteudo = parte.get_payload(decode=True).decode("utf-8", errors="ignore")
                                        codigo = extrair_codigo(conteudo)
                                        if codigo:
                                            return codigo
                            else:
                                conteudo = msg.get_payload(decode=True).decode("utf-8", errors="ignore")
                                codigo = extrair_codigo(conteudo)
                                if codigo:
                                    return codigo
        print("Nenhum código encontrado nos e-mails.")
        return None

    except Exception as e:
        print(f"Erro ao buscar o código no e-mail: {e}")
        return None

# Configuração principal
nome_arquivo = "login_imap.txt"
usuario_portal, senha_portal = carregar_credenciais(nome_arquivo)

if not usuario_portal or not senha_portal:
    print("Erro ao carregar credenciais. Encerrando...")
else:
    assunto_filtro = "Código de verificação - Login: jvsreis"  # Assunto esperado no e-mail
    servidor = conectar_email(usuario_portal, senha_portal)

    if servidor:
        print("Buscando código de verificação...")
        codigo_6 = buscar_codigo_por_assunto(servidor, assunto_filtro)

        if codigo_6:
            print(f"Código encontrado: {codigo_6}")
            # Preencher o campo com o código coletado
            xpath_campo_codigo = '//*[@id="codigo"]'  # Substitua pelo XPath do campo desejado
            preencher_campo_xpath(navegador, xpath_campo_codigo, codigo_6)
        else:
            print("Nenhum código encontrado.")

        servidor.logout()


# inserir codigo coletado no email
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="codigo"]')
if elemento_senha_portal:
    elemento_senha_portal.send_keys(codigo_6)
else:
    print("Elemento 'Acessar' não encontrado!")

# Código principal
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="acessar1"]') #acessar final
if elemento_senha_portal:
    elemento_senha_portal.click()
    # Verifica se tem alguma notificação de alerta
    verificar_e_clicar(navegador, '//*[@id="boxes"]/input')
else:
    print("Elemento 'Acessar' não encontrado!")





# (Pagadores)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="format_layout_content"]/div[2]/div[1]/div[2]/div/div[2]/div[1]/div/div/div/div/ul/li[2]/a')
if elemento_senha_portal:
    elemento_senha_portal.click()
else:
    print("Elemento 'Usuários Pagadores' não encontrado!")

# (Importar 1)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="menu_content"]/div[4]/div/div[2]/input[3]')
if elemento_senha_portal:
    elemento_senha_portal.click()
else:
    print("Elemento 'Importar 1' não encontrado!")
time.sleep(2)

#################################### Automação De Unidades #################################################

pyautogui.PAUSE = 0.3  # Pausa padrão entre comandos
pyautogui.FAILSAFE = True  # Move o mouse para o canto superior esquerdo para interromper

def move_and_click(x, y, duration=0.1):
    """Move o mouse para uma posição e clica."""
    pyautogui.moveTo(x, y, duration=duration)
    pyautogui.click()

def double_click(x, y, duration=0.1):
    """Move o mouse para uma posição e dá um duplo clique."""
    pyautogui.moveTo(x, y, duration=duration)
    pyautogui.doubleClick()

def press_key(*keys, duration=0.1):
    """Pressiona uma sequência de teclas."""
    for key in keys:
        pyautogui.press(key)
        time.sleep(duration)







































































#####################################################################################
#####################################################################################
#####################################################################################

# -------------------------- Usuarios_pagadores.csv ---------------------------------

# Usuarios_pagadores.csv
navegador.get("https://www.portaldeboletos.com.br/brmarinas?ctr=usuariosPagador&mt=index")

# (Importar usuarios pagadores)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="menu_content"]/div[4]/div/div[2]/input[3]')
if elemento_senha_portal:
    elemento_senha_portal.click()
else:
    print("Elemento 'Importar 1' não encontrado!")


def limpar_pasta(caminho_pasta):
    """
    Remove todos os arquivos da pasta especificada.
    :param caminho_pasta: str - Caminho da pasta a ser limpa
    """
    try:
        for arquivo in os.listdir(caminho_pasta):
            nome_arquivo = os.path.join(caminho_pasta, arquivo)
            if os.path.isfile(nome_arquivo):
                os.remove(nome_arquivo)
                print(f"Arquivo removido: {nome_arquivo}")
        print("Pasta limpa com sucesso.")
    except Exception as e:
        print(f"Erro ao limpar a pasta: {e}")
        return False
    return True

time.sleep(1)  
navegador.quit()