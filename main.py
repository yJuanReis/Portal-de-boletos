from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from email.header import decode_header
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

import time
import email
import imaplib
import os
import re
import sys
import locale
import pyautogui

escolher_arquivo = 'escolher_arquivo.png'
NOME_ARQUIVO_CONFIG = "login_imap.txt"

def limpar_pasta(caminho_pasta):
    """
    Remove todos os arquivos da pasta especificada.
    :param caminho_pasta: str - Caminho da pasta a ser limpa
    """
    try:
        for arquivo in os.listdir(caminho_pasta):
            caminho_arquivo = os.path.join(caminho_pasta, arquivo)
            if os.path.isfile(caminho_arquivo):
                os.remove(caminho_arquivo)
                print(f"Arquivo removido: {caminho_arquivo}")
        print("Pasta limpa com sucesso.")
    except Exception as e:
        print(f"Erro ao limpar a pasta: {e}")

def obter_caminho_pasta(nome_arquivo):
    """Obtém o caminho da pasta especificada na linha 5 do arquivo."""
    try:
        with open(nome_arquivo, 'r', encoding='utf-8') as arquivo:
            linhas = arquivo.readlines()
            if len(linhas) >= 5:
                return linhas[4].strip()
            else:
                raise ValueError("O arquivo não contém 5 linhas.")
    except Exception as e:
        print(f"Erro ao ler o arquivo {nome_arquivo}: {e}")
        return None

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
        with open(nome_arquivo, "r", encoding="utf-8") as arquivo:
            linhas = arquivo.readlines()
            usuario_portal = linhas[0].strip()  # Primeira linha: e-mail
            senha_portal = linhas[1].strip()   # Segunda linha: senha
            print("Credenciais carregadas com sucesso.")
            return usuario_portal, senha_portal
    except Exception as e:
        print(f"Erro ao carregar credenciais do arquivo {nome_arquivo}: {e}")
        return None, None

# Configurações do script
nome_arquivo_config = "login_imap.txt"
usuario_portal, senha_portal = carregar_credenciais(nome_arquivo_config)

if not usuario_portal or not senha_portal:
    print("Erro ao carregar credenciais. Encerrando...")
else:
    nome_arquivo_procurado = "usuarios_pagadores.csv"
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
        campo = WebDriverWait(navegador, 10).until(  # Increased wait time from 5 to 10 seconds
            EC.presence_of_element_located((By.XPATH, xpath_campo))
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

# Obtendo o caminho da pasta de downloads a partir do arquivo de configuração
caminho_pasta_downloads = obter_caminho_pasta(NOME_ARQUIVO_CONFIG)

# Configurando e inicializando o navegador automaticamente
service = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=service)
navegador.maximize_window()
navegador.get("https://www.portaldeboletos.com.br/brmarinas")

with open(NOME_ARQUIVO_CONFIG, "r", encoding='utf-8') as arquivo:
    linhas = arquivo.readlines()
    usuario_portal = linhas[0].strip()
    senha_portal = linhas[1].strip()
    usuario_portal = linhas[2].strip()
    senha_portal = linhas[3].strip()
    pasta_destino = obter_caminho_pasta(NOME_ARQUIVO_CONFIG)

elemento_usuario_portal = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frm_tab_usuario_senha"]/div/input[2]')))
elemento_usuario_portal.send_keys(usuario_portal)

elemento_senha_portal = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frm_tab_usuario_senha"]/div/input[3]')))
elemento_senha_portal.send_keys(senha_portal)

elemento_senha_portal = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="acessar1"]')))
elemento_senha_portal.click()

time.sleep(3)

def carregar_credenciais(nome_arquivo):
    """Carrega as credenciais de e-mail do arquivo especificado."""
    try:
        with open(nome_arquivo, "r", encoding="utf-8") as arquivo:
            linhas = arquivo.readlines()
            usuario_portal = linhas[0].strip()  # Primeira linha: e-mail
            senha_portal = linhas[1].strip()  # Segunda linha: senha
            return usuario_portal, senha_portal
    except Exception as e:
        print(f"Erro ao carregar credenciais: {e}")
        return None, None

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
nome_arquivo_config = "login_imap.txt"
usuario_portal, senha_portal = carregar_credenciais(nome_arquivo_config)

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

time.sleep(1)
elemento_senha_portal = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="acessar1"]')))  # enviar
elemento_senha_portal.click()

time.sleep(1)
elemento_senha_portal = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="format_layout_content"]/div[2]/div[1]/div[2]/div/div[2]/div[1]/div/div/div/div/ul/li[3]/a')))  # usuarios pagadores
elemento_senha_portal.click()

time.sleep(1)
elemento_senha_portal = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="menu_content"]/div[4]/div/div[2]/input[3]')))  # importar 1
elemento_senha_portal.click()
time.sleep(5)


############################################################################################################
#################################### Automação De Unidades #################################################
############################################################################################################


# Imagens para o clique
escolher_arquivo = 'escolher_arquivo.png'
downloads = 'downloads.png'
usuarios_pagadores = 'usuarios_pagadores.png'

def click_coordinates(x, y):
    pyautogui.moveTo(x, y)
    pyautogui.click()
    time.sleep(delay)

delay = 1  # Em segundos

coordinates = [
    (620, 252),   # Unidade de negócio
    (623, 304),   # Verolme
    (980, 248),   # SELECIONAR PAGADOR 
    (980, 306)    # PAGADOR 
]

for coord in coordinates:
    click_coordinates(coord[0], coord[1])  # Clica na coordenada atual
    
    # Clica no botão de upload via Selenium após cada clique nas coordenadas
    try:
     
        time.sleep(delay)
    except Exception as e:
        print(f"Erro ao clicar no botão de upload: {e}")
        # navegador.quit()

try:
   
    x, y = pyautogui.locateCenterOnScreen(escolher_arquivo, confidence=0.8)  # Ajuste 'confidence' se necessário
    print(f"Centro da imagem encontrado em: ({x}, {y})")

    # Realiza o clique no centro da imagem
    pyautogui.click(x, y)
    print("Clique realizado com sucesso!")
except TypeError:
    print(f"Erro ao clicar no botão de upload: {e}")
    # navegador.quit()

time.sleep(1)

area_downloads = (0, 0, 800, 600)  # Ajuste conforme necessário

try:
    # Localiza a imagem na região especificada com confiança de 80%
    imagem_localizada = pyautogui.locateOnScreen(downloads, region=area_downloads, confidence=0.8,delay=1)

    if imagem_localizada:
        # Obtém o centro da área onde a imagem foi encontrada
        centro_x, centro_y = pyautogui.center(imagem_localizada)
        
        # Realiza o clique no centro da imagem
        pyautogui.click(centro_x, centro_y)
        print(f"Imagem encontrada e clicada em ({centro_x}, {centro_y})")
    else:
        print("Imagem não encontrada na região especificada.")
except Exception as e:
    print(f"Ocorreu um erro: {e}")

time.sleep(1)

try:
   
    x, y = pyautogui.locateCenterOnScreen(usuarios_pagadores, confidence=0.6)  # Ajuste 'confidence' se necessário
    print(f"Centro da imagem encontrado em: ({x}, {y})")

    # Realiza o clique no centro da imagem
    pyautogui.doubleClick(x, y)
    print("Clique realizado com sucesso!")
except TypeError:
    print(f"Erro ao clicar no botão de upload: {e}")
    # navegador.quit()


elemento_senha_portal = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="enviar_frm_upload"]')))  # importar 2
elemento_senha_portal.click()

elemento_senha_portal = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="msgConfirmContent"]/div/div[2]/div[3]/input[2]')))  # OK
elemento_senha_portal.click()

elemento_senha_portal = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="msgSuccessContent"]/div/div[2]/div[3]/input')))  # Fechar
elemento_senha_portal.click()

##################################################################################################################################################





time.sleep(100000)  
navegador.quit()













