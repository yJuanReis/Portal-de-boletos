import keyboard
from selenium import webdriver
import sys
import locale
import time
import email
import imaplib
import re
import pyautogui
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from email.header import decode_header
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

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
nome_arquivo = 'login_imap.txt'

# Busca o arquivo no diretório do script e subdiretórios
caminho_arquivo = encontrar_arquivo(nome_arquivo, script_dir)

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
        # Obtém o diretório onde o script está localizado
        script_dir = os.path.dirname(os.path.abspath(__file__))
        caminho_arquivo = os.path.join(script_dir, nome_arquivo)

        # Abre o arquivo no diretório do script
        with open(caminho_arquivo, 'r', encoding='utf-8') as arquivo:
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
        caminho_arquivo = os.path.join(script_dir, nome_arquivo)

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_arquivo):
            raise FileNotFoundError(f"O arquivo {nome_arquivo} não foi encontrado no diretório do script.")

        # Lê o arquivo e carrega as credenciais
        with open(caminho_arquivo, "r", encoding="utf-8") as arquivo:
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
caminho_arquivo = "login_imap.txt"
usuario_portal, senha_portal = carregar_credenciais(caminho_arquivo)

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

# Configurando e inicializando o navegador automaticamente
service = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=service)
navegador.maximize_window()
navegador.get("https://www.portaldeboletos.com.br/brmarinas")

with open(caminho_arquivo, "r", encoding='utf-8') as arquivo:
    linhas = arquivo.readlines()
    email_imap = linhas[0].strip()
    senha_app = linhas[1].strip()
    usuario_portal = linhas[2].strip()
    senha_portal = linhas[3].strip()
    pasta_destino = obter_caminho_pasta(caminho_arquivo)

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
usuario_portal, senha_portal = carregar_credenciais(caminho_arquivo)

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

# -------------------------- verolme --------------------------
# Etapa 1: Seleciona a Unidade de negocio
move_and_click(900, 390)  # Clica no centro da tela
time.sleep(fast)
press_key('tab')         
time.sleep(fast)

#  1 - Verolme
for _ in range(1):
    press_key('down')

# Etapa 2: Escolher arquivo
move_and_click(644, 335) 
time.sleep(delay)
pyautogui.hotkey('ctrl', 'l')  # Barra de Caminho
pyautogui.write(caminho_pasta)  
press_key('enter')            # Confirma

# Etapa 3: Seleção do arquivo e confirmação
move_and_click(537, 299) 
time.sleep(delay)
press_key('down')             # Ajusta a seleção com Up e Down (opcional)
press_key('enter')            # Confirma a seleção do arquivo com Enter
time.sleep(delay)
# Etapa 4: - Importar
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar
time.sleep(tempo)

# Etapa 5: Fechar
move_and_click(900, 583)       # Move e clica no botão Fechar final
time.sleep(fast)
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar



# -------------------------- Gloria --------------------------------
# Etapa 1: Seleciona a Unidade de negocio
move_and_click(900, 390)  # Clica no centro da tela
time.sleep(fast)
press_key('tab')         
time.sleep(fast)

# 2 - Gloria 
for _ in range(2):
    press_key('down')

# Etapa 2: Escolher arquivo - Seleção do arquivo e confirmação
move_and_click(644, 335) 
move_and_click(537, 299) 
time.sleep(delay)
press_key('down')             # Ajusta a seleção com Up e Down (opcional)
press_key('enter')            # Confirma a seleção do arquivo com Enter
time.sleep(delay)
# Etapa 3: - Importar
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar
time.sleep(tempo)

# Etapa 4: Fechar
move_and_click(900, 583)       # Move e clica no botão Fechar final
time.sleep(fast)
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar

# --------------------------------------------------------------------

# -------------------------- Itacuruça --------------------------
# Etapa 1: Seleciona a Unidade de negocio
move_and_click(900, 390)  # Clica no centro da tela
time.sleep(fast)
press_key('tab')         
time.sleep(fast)

# 3 - Itacuruça 
for _ in range(3):
    press_key('down')

# Etapa 2: Escolher arquivo - Seleção do arquivo e confirmação
move_and_click(644, 335) 
move_and_click(537, 299) 
time.sleep(delay)
press_key('down')             # Ajusta a seleção com Up e Down (opcional)
press_key('enter')            # Confirma a seleção do arquivo com Enter
time.sleep(delay)
# Etapa 3: - Importar
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar
time.sleep(tempo)

# Etapa 4: Fechar
move_and_click(900, 583)       # Move e clica no botão Fechar final
time.sleep(fast)
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar

# --------------------------------------------------------------------

# -------------------------- Piratas --------------------------

# Etapa 1: Seleciona a Unidade de negocio
move_and_click(900, 390)  # Clica no centro da tela
time.sleep(fast)
press_key('tab')         
time.sleep(fast)

# 4 - Piratas 
for _ in range(4):
    press_key('down')

# Etapa 2: Escolher arquivo - Seleção do arquivo e confirmação
move_and_click(644, 335) 
move_and_click(537, 299) 
time.sleep(delay)
press_key('down')             # Ajusta a seleção com Up e Down (opcional)
press_key('enter')            # Confirma a seleção do arquivo com Enter
time.sleep(delay)
# Etapa 3: - Importar
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar
time.sleep(tempo)

# Etapa 4: Fechar
move_and_click(900, 583)       # Move e clica no botão Fechar final
time.sleep(fast)
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar

# --------------------------------------------------------------------

# -------------------------- Porto Bracuhy ---------------------------

# Etapa 1: Seleciona a Unidade de negocio
move_and_click(900, 390)  # Clica no centro da tela
time.sleep(fast)
press_key('tab')         
time.sleep(fast)

# 6 - Porto Bracuhy  
for _ in range(6):
    press_key('down')

# Etapa 2: Escolher arquivo - Seleção do arquivo e confirmação
move_and_click(644, 335) 
move_and_click(537, 299) 
time.sleep(delay)
press_key('down')             # Ajusta a seleção com Up e Down (opcional)
press_key('enter')            # Confirma a seleção do arquivo com Enter
time.sleep(delay)
# Etapa 3: - Importar
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar
time.sleep(tempo)

# Etapa 4: Fechar
move_and_click(900, 583)       # Move e clica no botão Fechar final
time.sleep(fast)
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar

# --------------------------------------------------------------------

# -------------------------- Refugio de Paraty --------------------------

# Etapa 1: Seleciona a Unidade de negocio
move_and_click(900, 390)  # Clica no centro da tela
time.sleep(fast)
press_key('tab')         
time.sleep(fast)

# 7 - Refugio de Paraty
for _ in range(7):
    press_key('down')

# Etapa 2: Escolher arquivo - Seleção do arquivo e confirmação
move_and_click(644, 335) 
move_and_click(537, 299) 
time.sleep(delay)
press_key('down')             # Ajusta a seleção com Up e Down (opcional)
press_key('enter')            # Confirma a seleção do arquivo com Enter
time.sleep(delay)
# Etapa 3: - Importar
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar
time.sleep(tempo)

# Etapa 4: Fechar
move_and_click(900, 583)       # Move e clica no botão Fechar final
time.sleep(fast)
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar

# --------------------------------------------------------------------

# -------------------------- Ribeira --------------------------

# Etapa 1: Seleciona a Unidade de negocio
move_and_click(900, 390)  # Clica no centro da tela
time.sleep(fast)
press_key('tab')         
time.sleep(fast)

# 8 - Ribeira
for _ in range(8):
    press_key('down')

# Etapa 2: Escolher arquivo - Seleção do arquivo e confirmação
move_and_click(644, 335) 
move_and_click(537, 299) 
time.sleep(delay)
press_key('down')             # Ajusta a seleção com Up e Down (opcional)
press_key('enter')            # Confirma a seleção do arquivo com Enter
time.sleep(delay)
# Etapa 3: - Importar
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar
time.sleep(tempo)

# Etapa 4: Fechar
move_and_click(900, 583)       # Move e clica no botão Fechar final
time.sleep(fast)
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar

# --------------------------------------------------------------------

# -------------------------- Buzios --------------------------

# Etapa 1: Seleciona a Unidade de negocio
move_and_click(900, 390)  # Clica no centro da tela
time.sleep(fast)
press_key('tab')         
time.sleep(fast)

# 9 - Buzios
for _ in range(9):
    press_key('down')

# Etapa 2: Escolher arquivo - Seleção do arquivo e confirmação
move_and_click(644, 335) 
move_and_click(537, 299) 
time.sleep(delay)
press_key('down')             # Ajusta a seleção com Up e Down (opcional)
press_key('enter')            # Confirma a seleção do arquivo com Enter
time.sleep(delay)
# Etapa 3: - Importar
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar
time.sleep(tempo)

# Etapa 4: Fechar
move_and_click(900, 583)       # Move e clica no botão Fechar final
time.sleep(fast)
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar

# --------------------------------------------------------------------

# -------------------------- Estacinamento --------------------------

# Etapa 1: Seleciona a Unidade de negocio
move_and_click(900, 390)  # Clica no centro da tela
time.sleep(fast)
press_key('tab')         
time.sleep(fast)

# 10 - Estacinamento
for _ in range(10):
    press_key('down')

# Etapa 2: Escolher arquivo - Seleção do arquivo e confirmação
move_and_click(644, 335) 
move_and_click(537, 299) 
time.sleep(delay)
press_key('down')             # Ajusta a seleção com Up e Down (opcional)
press_key('enter')            # Confirma a seleção do arquivo com Enter
time.sleep(delay)
# Etapa 3: - Importar
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar
time.sleep(tempo)

# Etapa 4: Fechar
move_and_click(900, 583)       # Move e clica no botão Fechar final
time.sleep(fast)
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar

# --------------------------------------------------------------------


# -------------------------- JL Bracuhy --------------------------

# Etapa 1: Seleciona a Unidade de negocio
move_and_click(900, 390)  # Clica no centro da tela
time.sleep(fast)
press_key('tab')         
time.sleep(fast)

# 11 - JL Bracuhy
for _ in range(11):
    press_key('down')

# Etapa 2: Escolher arquivo - Seleção do arquivo e confirmação
move_and_click(644, 335) 
move_and_click(537, 299) 
time.sleep(delay)
press_key('down')             # Ajusta a seleção com Up e Down (opcional)
press_key('enter')            # Confirma a seleção do arquivo com Enter
time.sleep(delay)
# Etapa 3: - Importar
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar
time.sleep(tempo)

# Etapa 4: Fechar
move_and_click(900, 583)       # Move e clica no botão Fechar final
time.sleep(fast)
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar

# --------------------------------------------------------------------


# -------------------------- Boa Vista --------------------------

# Etapa 1: Seleciona a Unidade de negocio
move_and_click(900, 390)  # Clica no centro da tela
time.sleep(fast)
press_key('tab')         
time.sleep(fast)

# 12 - Boa Vista 
for _ in range(12):
    press_key('down')

# Etapa 2: Escolher arquivo - Seleção do arquivo e confirmação
move_and_click(644, 335) 
move_and_click(537, 299) 
time.sleep(delay)
press_key('down')             # Ajusta a seleção com Up e Down (opcional)
press_key('enter')            # Confirma a seleção do arquivo com Enter
time.sleep(delay)
# Etapa 3: - Importar
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar
time.sleep(tempo)

# Etapa 4: Fechar
move_and_click(900, 583)       # Move e clica no botão Fechar final
time.sleep(fast)
pyautogui.hotkey('tab', 'enter')  # Atalho para finalizar


#--------------------------------- Automação Usuarios Pagadores-------------------------

# Usuarios_pagadores.csv
navegador.get("https://www.portaldeboletos.com.br/brmarinas?ctr=usuariosPagador&mt=index")

# (Importar usuarios pagadores)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="menu_content"]/div[4]/div/div[2]/input[3]')
if elemento_senha_portal:
    elemento_senha_portal.click()
else:
    print("Elemento 'Importar 1' não encontrado!")

# -------------------------- Usuarios_pagadores.csv ---------------------------------

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
        # Obtém o diretório onde o script está localizado
        script_dir = os.path.dirname(os.path.abspath(__file__))
        caminho_arquivo = os.path.join(script_dir, nome_arquivo)

        # Abre o arquivo no diretório do script
        with open(caminho_arquivo, 'r', encoding='utf-8') as arquivo:
            linhas = arquivo.readlines()
            if len(linhas) >= 5:
                return linhas[4].strip()
            else:
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
        caminho_arquivo = os.path.join(script_dir, nome_arquivo)

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_arquivo):
            raise FileNotFoundError(f"O arquivo {nome_arquivo} não foi encontrado no diretório do script.")

        # Lê o arquivo e carrega as credenciais
        with open(caminho_arquivo, "r", encoding="utf-8") as arquivo:
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
caminho_arquivo = "login_imap.txt"
usuario_portal, senha_portal = carregar_credenciais(caminho_arquivo)
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

# 5m para começar a importação do Usuarios

time.sleep(600)
############################################################################################################
################################# Automação De Unidades usuarios pagadores #################################
############################################################################################################

# -------------------------- verolme --------------------------
pyautogui.moveTo(900, 390)
pyautogui.click()
time.sleep(fast)
press_key('tab')
time.sleep(delay)

#1 verolme
for _ in range(1):
    press_key('down')
#------------------
press_key('tab')
time.sleep(delay)
press_key('down')
pyautogui.moveTo(827, 352)
pyautogui.click()
time.sleep(tempo)


#Escreve o caminho do downloads na barra de caminho e da enter
pyautogui.hotkey('ctrl', 'l') 
pyautogui.write(caminho_pasta)
press_key('enter')
for _ in range(5):
    pyautogui.hotkey('tab')
#--------------------------

#-----padrão ------
press_key('up')
press_key('down')
time.sleep(delay)
#Da Enter no arquivo selecionado
press_key('enter')
time.sleep(tempo)
#clica em importar final
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="enviar_frm_upload"]')
elemento_senha_portal.click()
time.sleep(tempo)
#Clica em OK
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="msgConfirmContent"]/div/div[2]/div[3]/input[2]')
elemento_senha_portal.click()
time.sleep(tempo)
#clica em Fechar final
pyautogui.moveTo(900, 583)
pyautogui.click()
time.sleep(delay)
pyautogui.hotkey('tab', 'enter')

# -------------------------- Gloria --------------------------
def press_key(key):
    pyautogui.press(key)
    time.sleep(fast)   
pyautogui.moveTo(900, 390)
pyautogui.click()
time.sleep(fast)
press_key('tab')
time.sleep(fast)

# 2 gloria
for _ in range(2):
    press_key('down')
#------------------
press_key('tab')
press_key('down')
time.sleep(delay)
pyautogui.moveTo(827, 352)
pyautogui.click()
time.sleep(fast)
pyautogui.moveTo(507, 252)
pyautogui.doubleClick()
time.sleep(fast)
press_key('down')
press_key('enter')
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="enviar_frm_upload"]')
elemento_senha_portal.click()
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="msgConfirmContent"]/div/div[2]/div[3]/input[2]')
elemento_senha_portal.click()
time.sleep(tempo)
pyautogui.moveTo(900, 583)
pyautogui.click()
time.sleep(delay)
pyautogui.hotkey('tab', 'enter')

# -------------------------- Itacuruça --------------------------
pyautogui.moveTo(900, 390)
pyautogui.click()
time.sleep(fast)
press_key('tab')
time.sleep(delay)

# 3 itacuruça
for _ in range(3):
    press_key('down')

#------------------
press_key('tab')
press_key('down')
time.sleep(delay)
pyautogui.moveTo(827, 352)
pyautogui.click()
time.sleep(fast)
pyautogui.moveTo(507, 252)
pyautogui.doubleClick()
time.sleep(fast)
press_key('down')
press_key('enter')
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="enviar_frm_upload"]')
elemento_senha_portal.click()
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="msgConfirmContent"]/div/div[2]/div[3]/input[2]')
elemento_senha_portal.click()
time.sleep(tempo)
pyautogui.moveTo(900, 583)
pyautogui.click()
time.sleep(delay)
pyautogui.hotkey('tab', 'enter')

# -------------------------- Piratas --------------------------
pyautogui.moveTo(900, 390)
pyautogui.click()
time.sleep(fast)
press_key('tab')
time.sleep(delay)

# 4 piratas
for _ in range(4):
    press_key('down')

#------------------
press_key('tab')
press_key('down')
time.sleep(delay)
pyautogui.moveTo(827, 352)
pyautogui.click()
time.sleep(fast)
pyautogui.moveTo(507, 252)
pyautogui.doubleClick()
time.sleep(fast)
press_key('down')
press_key('enter')
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="enviar_frm_upload"]')
elemento_senha_portal.click()
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="msgConfirmContent"]/div/div[2]/div[3]/input[2]')
elemento_senha_portal.click()
time.sleep(tempo)
pyautogui.moveTo(900, 583)
pyautogui.click()
time.sleep(delay)
pyautogui.hotkey('tab', 'enter')

# ------------------------ Porto Bracuhy ----------------------
pyautogui.moveTo(900, 390)
pyautogui.click()
time.sleep(fast)
press_key('tab')
time.sleep(delay)

# 6
for _ in range(6):
    press_key('down')

#------------------
press_key('tab')
press_key('down')
time.sleep(delay)
pyautogui.moveTo(827, 352)
pyautogui.click()
time.sleep(fast)
pyautogui.moveTo(507, 252)
pyautogui.doubleClick()
time.sleep(fast)
press_key('down')
press_key('enter')
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="enviar_frm_upload"]')
elemento_senha_portal.click()
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="msgConfirmContent"]/div/div[2]/div[3]/input[2]')
elemento_senha_portal.click()
time.sleep(tempo)
pyautogui.moveTo(900, 583)
pyautogui.click()
time.sleep(delay)
pyautogui.hotkey('tab', 'enter')

# -------------------------- Refugio de Paraty --------------------------
pyautogui.moveTo(900, 390)
pyautogui.click()
time.sleep(fast)
press_key('tab')
time.sleep(delay)

# 7
for _ in range(7):
    press_key('down')

#------------------
press_key('tab')
press_key('down')
time.sleep(delay)
pyautogui.moveTo(827, 352)
pyautogui.click()
time.sleep(fast)
pyautogui.moveTo(507, 252)
pyautogui.doubleClick()
time.sleep(fast)
press_key('down')
press_key('enter')
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="enviar_frm_upload"]')
elemento_senha_portal.click()
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="msgConfirmContent"]/div/div[2]/div[3]/input[2]')
elemento_senha_portal.click()
time.sleep(tempo)
pyautogui.moveTo(900, 583)
pyautogui.click()
time.sleep(delay)
pyautogui.hotkey('tab', 'enter')

# -------------------------- Ribeira --------------------------
pyautogui.moveTo(900, 390)
pyautogui.click()
time.sleep(fast)
press_key('tab')
time.sleep(delay)

# 8
for _ in range(8):
    press_key('down')

#------------------
press_key('tab')
press_key('down')
time.sleep(delay)
pyautogui.moveTo(827, 352)
pyautogui.click()
time.sleep(fast)
pyautogui.moveTo(507, 252)
pyautogui.doubleClick()
time.sleep(fast)
press_key('down')
press_key('enter')
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="enviar_frm_upload"]')
elemento_senha_portal.click()
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="msgConfirmContent"]/div/div[2]/div[3]/input[2]')
elemento_senha_portal.click()
time.sleep(tempo)
pyautogui.moveTo(900, 583)
pyautogui.click()
time.sleep(delay)
pyautogui.hotkey('tab', 'enter')

# -------------------------- Buzios --------------------------
pyautogui.moveTo(900, 390)
pyautogui.click()
time.sleep(fast)
press_key('tab')
time.sleep(delay)

# 9 
for _ in range(9):
    press_key('down')

#------------------
press_key('tab')
press_key('down')
time.sleep(delay)
pyautogui.moveTo(827, 352)
pyautogui.click()
time.sleep(fast)
pyautogui.moveTo(507, 252)
pyautogui.doubleClick()
time.sleep(fast)
press_key('down')
press_key('enter')
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="enviar_frm_upload"]')
elemento_senha_portal.click()
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="msgConfirmContent"]/div/div[2]/div[3]/input[2]')
elemento_senha_portal.click()
time.sleep(tempo)
pyautogui.moveTo(900, 583)
pyautogui.click()
time.sleep(delay)
pyautogui.hotkey('tab', 'enter')

# -------------------------- Estacionamento --------------------------
pyautogui.moveTo(900, 390)
pyautogui.click()
time.sleep(fast)
press_key('tab')
time.sleep(delay)

# 10
for _ in range(10):
    press_key('down')

press_key('tab')
press_key('down')
time.sleep(delay)
#------------------
pyautogui.moveTo(827, 352)
pyautogui.click()
time.sleep(fast)
pyautogui.moveTo(507, 252)
pyautogui.doubleClick()
time.sleep(fast)
press_key('down')
press_key('enter')
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="enviar_frm_upload"]')
elemento_senha_portal.click()
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="msgConfirmContent"]/div/div[2]/div[3]/input[2]')
elemento_senha_portal.click()
time.sleep(tempo)
pyautogui.moveTo(900, 583)
pyautogui.click()
time.sleep(delay)
pyautogui.hotkey('tab', 'enter')

# -------------------------- JL Bracuhy --------------------------
pyautogui.moveTo(900, 390)
pyautogui.click()
time.sleep(fast)
press_key('tab')
time.sleep(delay)

# 11
for _ in range(11):
    press_key('down')

#------------------
press_key('tab')
press_key('down')
time.sleep(delay)
pyautogui.moveTo(827, 352)
pyautogui.click()
time.sleep(fast)
pyautogui.moveTo(507, 252)
pyautogui.doubleClick()
time.sleep(fast)
press_key('down')
press_key('enter')
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="enviar_frm_upload"]')
elemento_senha_portal.click()
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="msgConfirmContent"]/div/div[2]/div[3]/input[2]')
elemento_senha_portal.click()
time.sleep(delay)
pyautogui.moveTo(900, 583)
pyautogui.click()
time.sleep(delay)
pyautogui.hotkey('tab', 'enter')

# -------------------------- Boa vista --------------------------

pyautogui.moveTo(900, 390)
pyautogui.click()
time.sleep(fast)
press_key('tab')
time.sleep(fast)

# 12
for _ in range(12):
    press_key('down')

#------------------
press_key('tab')
press_key('down')
time.sleep(delay)

pyautogui.moveTo(827, 352)
pyautogui.click()
time.sleep(fast)
pyautogui.moveTo(507, 252)
pyautogui.doubleClick()
time.sleep(fast)
press_key('down')
press_key('enter')
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="enviar_frm_upload"]')
elemento_senha_portal.click()
time.sleep(tempo)
elemento_senha_portal = esperar_elemento(navegador, '//*[@id="msgConfirmContent"]/div/div[2]/div[3]/input[2]')
elemento_senha_portal.click()
time.sleep(tempo)
pyautogui.moveTo(900, 583)
pyautogui.click()
time.sleep(delay)
pyautogui.hotkey('tab', 'enter')

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
        return False
    return True

time.sleep(1)  
navegador.quit()
