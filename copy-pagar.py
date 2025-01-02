from email.header import decode_header
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import os
import sys
import locale
import imaplib
import email
import pyautogui
import time
import re
import keyboard

# Configurações globais
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '2'
sys.stdout.reconfigure(encoding='utf-8')
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

####################################################################################################

# Nome do arquivo de configuração
NOME_ARQUIVO_CONFIG = "login_imap.txt"


# Funções auxiliares para manipulação de arquivos e pastas
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

# Função para limpar a pasta de downloads
def limpar_pasta_downloads(caminho_pasta):
    """Remove todos os arquivos da pasta especificada."""
    try:
        for arquivo in os.listdir(caminho_pasta):
            caminho_arquivo = os.path.join(caminho_pasta, arquivo)
            if os.path.isfile(caminho_arquivo):
                os.remove(caminho_arquivo)
                print(f"Arquivo removido: {caminho_arquivo}")
        print("Pasta limpa com sucesso.")
    except Exception as e:
        print(f"Erro ao limpar a pasta: {e}")

# Executar a limpeza da pasta assim que o código iniciar
if __name__ == "__main__":
    # Obter o caminho da pasta de downloads
    caminho_pasta_downloads = obter_caminho_pasta(NOME_ARQUIVO_CONFIG)

    # Verificar se o caminho é válido e limpar a pasta
    if caminho_pasta_downloads and os.path.isdir(caminho_pasta_downloads):
        print("Limpando a pasta de downloads...")
        limpar_pasta_downloads(caminho_pasta_downloads)
    else:
        print("Caminho da pasta inválido ou não encontrado.")

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
def baixar_anexo_por_email(usuario_email, senha_email, nome_arquivo_procurado, pasta_destino):
    """Busca e baixa um anexo específico de um e-mail."""
    try:
        # Conectar ao servidor IMAP
        servidor = imaplib.IMAP4_SSL("imap.gmail.com")
        servidor.login(usuario_email, senha_email)
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
            usuario_email = linhas[0].strip()  # Primeira linha: e-mail
            senha_email = linhas[1].strip()   # Segunda linha: senha
            print("Credenciais carregadas com sucesso.")
            return usuario_email, senha_email
    except Exception as e:
        print(f"Erro ao carregar credenciais do arquivo {nome_arquivo}: {e}")
        return None, None

# Configurações do script
nome_arquivo_config = "login_imap.txt"
usuario_email, senha_email = carregar_credenciais(nome_arquivo_config)

if not usuario_email or not senha_email:
    print("Erro ao carregar credenciais. Encerrando...")
else:
    nome_arquivo_procurado = "usuarios_pagadores.csv"
    pasta_destino = os.path.expanduser("~/Downloads")  # Altere conforme necessário

    # Executar a função para baixar o anexo
    print("Buscando e baixando o anexo...")
    sucesso = baixar_anexo_por_email(usuario_email, senha_email, nome_arquivo_procurado, pasta_destino)
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
def conectar_email(usuario_email, senha_email):
    """Conecta ao servidor de e-mail IMAP e retorna a conexão."""
    try:
        servidor = imaplib.IMAP4_SSL("imap.gmail.com")
        servidor.login(usuario_email, senha_email)
        return servidor
    except Exception as e:
        print(f"Erro ao conectar ao servidor de e-mail: {e}")
        return None

# Função para buscar o código de confirmação no e-mail
def buscar_codigo_confirmacao(servidor, remetente_filtro):
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

def configurar_navegador():
    """Configura o navegador Chrome usando o WebDriver Manager."""
    try:
        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico)
        navegador.maximize_window()  # Maximiza a janela do navegador
        return navegador
    except Exception as e:
        print(f"Erro ao configurar o navegador: {e}")
        return None
    


url = "https://www.portaldeboletos.com.br/brmarinas"  # Altere para o site que deseja abrir


def abrir_site(url):
    """Abre o site especificado na URL."""
    try:
        navegador = configurar_navegador()
        if navegador:
            navegador.get(url)
            print(f"Abrindo o site: {url}")
        else:
            print("Falha ao abrir o navegador.")
    except Exception as e:
        print(f"Erro ao abrir o site: {e}")

# Ler credenciais do arquivo de configuração
def credenciais():
    try:
        with open(NOME_ARQUIVO_CONFIG, "r") as arquivo:
            linhas = arquivo.readlines()
            usuario_email = linhas[0].strip()
            senha_email = linhas[1].strip()
            usuario_portal = linhas[2].strip()
            senha_portal = linhas[3].strip()
            pasta_destino = obter_caminho_pasta(NOME_ARQUIVO_CONFIG)
        
        print("Credenciais carregadas com sucesso.")
    except Exception as e:
        print(f"Erro ao carregar credenciais: {e}")
        return

    # Limpar a pasta de downloads (se aplicável)
    if pasta_destino and os.path.isdir(pasta_destino):
        limpar_pasta_downloads(pasta_destino)

    # Conectar ao servidor IMAP e buscar o código de confirmação
    servidor_imap = conectar_email(usuario_email, senha_email)
    if not servidor_imap:
        return

    remetente_filtro = "nao-responda@portaldeboletos.com.br"
    codigo_confirmacao = buscar_codigo_confirmacao(servidor_imap, remetente_filtro)

    if not codigo_confirmacao:
        print("Código de confirmação não encontrado.")
        return

    # Automação com Selenium para login no portal
    navegador = configurar_navegador()
    if not navegador:
        print("Falha ao configurar o navegador.")
        return
    
    # Realizar login no portal
    try:
        navegador.get("https://www.portaldeboletos.com.br/brmarinas")  # URL do portal
        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.NAME, "usuario"))).send_keys(usuario_portal)
        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.NAME, "senha"))).send_keys(senha_portal)
        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.NAME, "codigo_confirmacao"))).send_keys(codigo_confirmacao)
        WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.NAME, "entrar"))).click()
        print("Login realizado com sucesso.")
    except Exception as e:
        print(f"Erro ao realizar login no portal: {e}")
    finally:
        navegador.quit()  # Fechar o navegador após a operação