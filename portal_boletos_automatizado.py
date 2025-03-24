from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from webdriver_manager.firefox import GeckoDriverManager
from email.header import decode_header
import keyboard
import pyautogui
import imaplib
import email
import time
import os
import re
import sys
import locale

# Configurações globais
TEMPOS = {
    'delay': 0.7,
    'tempo': 1,
    'fast': 0.2,
    'wait': 2
}

class ConfiguracaoNavegador:
    @staticmethod
    def configurar_firefox():
        """Configura e retorna as opções do Firefox."""
        options = Options()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.binary_location = r"C:\Program Files\Mozilla Firefox\firefox.exe"
        return options

    @staticmethod
    def iniciar_navegador():
        """Inicia e retorna uma instância do navegador Firefox."""
        try:
            service = Service(GeckoDriverManager().install())
            options = ConfiguracaoNavegador.configurar_firefox()
            navegador = webdriver.Firefox(service=service, options=options)
            navegador.maximize_window()
            return navegador
        except Exception as e:
            print(f"Erro ao iniciar navegador: {e}")
            return None

class GerenciadorCredenciais:
    @staticmethod
    def carregar_credenciais(nome_arquivo):
        """Carrega credenciais do arquivo de configuração."""
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            caminho_arquivo = os.path.join(script_dir, nome_arquivo)
            
            with open(caminho_arquivo, "r", encoding="utf-8") as arquivo:
                linhas = arquivo.readlines()
                return {
                    'email_imap': linhas[0].strip(),
                    'senha_app': linhas[1].strip(),
                    'usuario_portal': linhas[2].strip(),
                    'senha_portal': linhas[3].strip(),
                    'caminho_pasta': linhas[4].strip() if len(linhas) > 4 else None
                }
        except Exception as e:
            print(f"Erro ao carregar credenciais: {e}")
            return None

class GerenciadorEmail:
    @staticmethod
    def conectar_email(email, senha):
        """Estabelece conexão com servidor de email."""
        try:
            servidor = imaplib.IMAP4_SSL("imap.gmail.com")
            servidor.login(email, senha)
            return servidor
        except Exception as e:
            print(f"Erro ao conectar ao email: {e}")
            return None

    @staticmethod
    def buscar_codigo_verificacao(servidor, assunto_filtro):
        """Busca código de verificação no email."""
        try:
            time.sleep(5)
            servidor.select("inbox")
            status, mensagens = servidor.search(None, "ALL")
            
            if status != "OK" or not mensagens[0]:
                return None

            for id_mensagem in reversed(mensagens[0].split()):
                status, msg_data = servidor.fetch(id_mensagem, "(RFC822)")
                if status == "OK":
                    email_msg = email.message_from_bytes(msg_data[0][1])
                    assunto = decode_header(email_msg["Subject"])[0][0]
                    if isinstance(assunto, bytes):
                        assunto = assunto.decode()
                    
                    if assunto_filtro in assunto:
                        for parte in email_msg.walk():
                            if parte.get_content_type() == "text/plain":
                                conteudo = parte.get_payload(decode=True).decode()
                                codigo = re.search(r"\b\d{6}\b", conteudo)
                                if codigo:
                                    return codigo.group(0)
            return None
        except Exception as e:
            print(f"Erro ao buscar código: {e}")
            return None

class AutomacaoPortal:
    def __init__(self):
        self.navegador = None
        self.credenciais = None
        print("\n=== Iniciando Automação do Portal de Boletos ===\n")

    def esperar_elemento(self, xpath, timeout=10):
        """Espera e retorna um elemento da página."""
        print(f"Procurando elemento: {xpath}")
        try:
            elemento = WebDriverWait(self.navegador, timeout).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            print("✓ Elemento encontrado com sucesso")
            return elemento
        except TimeoutException:
            print("✗ Elemento não encontrado")
            return None

    def fazer_login(self):
        """Realiza o processo de login no portal."""
        try:
            print("\n=== Iniciando Processo de Login ===")
            print("→ Acessando portal...")
            self.navegador.get("https://www.portaldeboletos.com.br/brmarinas")
            time.sleep(TEMPOS['delay'])
            
            print("\n→ Preenchendo credenciais...")
            # Preenche credenciais
            campos = {
                'usuario': '//*[@id="frm_tab_usuario_senha"]/div/input[2]',
                'senha': '//*[@id="frm_tab_usuario_senha"]/div/input[3]',
                'botao_acessar': '//*[@id="acessar1"]'
            }
            
            for campo, xpath in campos.items():
                print(f"  Preenchendo {campo}...")
                elemento = self.esperar_elemento(xpath)
                if elemento:
                    if campo in ['usuario', 'senha']:
                        elemento.send_keys(self.credenciais[f'{campo}_portal'])
                        print(f"  ✓ {campo.title()} preenchido")
                    else:
                        elemento.click()
                        print("  ✓ Botão acessar clicado")
                else:
                    raise Exception(f"Campo {campo} não encontrado")

            print("\n→ Processando código de verificação...")
            servidor = GerenciadorEmail.conectar_email(
                self.credenciais['email_imap'],
                self.credenciais['senha_app']
            )
            
            if servidor:
                print("  ✓ Conectado ao servidor de email")
                assunto = f"Código de verificação - Login: {self.credenciais['usuario_portal']}"
                print("  → Buscando código no email...")
                codigo = GerenciadorEmail.buscar_codigo_verificacao(servidor, assunto)
                
                if codigo:
                    print(f"  ✓ Código encontrado: {codigo}")
                    campo_codigo = self.esperar_elemento('//*[@id="codigo"]')
                    if campo_codigo:
                        campo_codigo.send_keys(codigo)
                        print("  ✓ Código inserido no campo")
                        
                        print("  → Confirmando código...")
                        self.esperar_elemento('//*[@id="acessar1"]').click()
                        print("  ✓ Código confirmado")
                    
                    print("\n→ Verificando notificações...")
                    try:
                        notif = self.esperar_elemento('//*[@id="boxes"]/input', 5)
                        if notif:
                            notif.click()
                            print("  ✓ Notificação fechada")
                    except:
                        print("  - Nenhuma notificação encontrada")
                
                servidor.logout()
                print("  ✓ Desconectado do servidor de email")
                return True
            
            return False
            
        except Exception as e:
            print(f"\n✗ Erro no processo de login: {e}")
            return False

    def acessar_importacao_usuarios(self):
        """Acessa a área de importação de usuários e seleciona a unidade."""
        try:
            print("\n=== Acessando Área de Importação ===")
            
            print("\n→ Navegando para área de usuários...")
            link_navegacao = self.esperar_elemento(
                "/html/body/div/div[1]/div[2]/div[1]/div[2]/div/div[2]/div[1]/div/div/div/div/ul/li[3]/a"
            )
            if link_navegacao:
                link_navegacao.click()
                time.sleep(TEMPOS['delay'])
                print("  ✓ Navegação realizada com sucesso")
            else:
                raise Exception("Link de navegação não encontrado")

            print("\n→ Acessando importação...")
            importar_btn = self.esperar_elemento(
                "/html/body/div/div[1]/div[2]/div/div[4]/div/div[2]/input[3]"
            )
            if importar_btn:
                importar_btn.click()
                time.sleep(TEMPOS['delay'])
                print("  ✓ Botão importar clicado")
            else:
                raise Exception("Botão importar não encontrado")

            print("\n→ Executando sequência de ações após importar...")
            
            # Clique no meio da tela (usando pyautogui)
            print("  → Clicando no meio da tela...")
            pyautogui.click(900, 390)  # Coordenadas do meio da tela
            time.sleep(TEMPOS['fast'])
            print("  ✓ Clique no meio da tela realizado")

            # Seleção da Unidade de Negócios
            print("  → Verolme...")
            sequencia_teclas = [
                ('tab'),
                ('down'), #Verolme

                ('tab') #Pagador
                ('down'),


                ('tab'), #Upload
                ('tab'),
                ('enter')
            ]

            for tecla, descricao in sequencia_teclas:
                print(f"    → Pressionando {descricao}...")
                pyautogui.press(tecla)
                time.sleep(TEMPOS['fast'])
                print(f"    ✓ {descricao} pressionado")

            # Executar Ctrl + L
            print("  → Executando Ctrl + L...")
            pyautogui.hotkey('ctrl', 'l')
            time.sleep(TEMPOS['fast'])
            print("  ✓ Ctrl + L executado")

            # Digitar o caminho do arquivo
            print("  → Digitando caminho do arquivo...")
            caminho = self.credenciais['caminho_pasta']
            pyautogui.write(caminho)
            time.sleep(TEMPOS['fast'])

            # Sequência de teclas após digitar o caminho
            print("  → Executando sequência de teclas após caminho...")
            pyautogui.press('enter')
            time.sleep(TEMPOS['fast'])
            
            print("  → Pressionando TAB 5 vezes com intervalo de 0.5 segundos...")
            for _ in range(5):
                pyautogui.press('tab')
                time.sleep(0.5)  # Aumentado para meio segundo
            print("  ✓ Sequência de TABs concluída")
            
            pyautogui.press('down')
            time.sleep(TEMPOS['fast'])
            pyautogui.press('enter')
            print("  ✓ Sequência de teclas executada com sucesso")

            # Clicando no botão de enviar
            print("  → Clicando em importar...")
            botao_enviar = self.esperar_elemento('//*[@id="enviar_frm_upload"]')
            if botao_enviar:
                botao_enviar.click()
                time.sleep(TEMPOS['wait'])  # Aumentado para 2 segundos
                print("  ✓ Botão de enviar clicado")

            # Clicando no botão OK
            print("  → Clicando no botão OK...")
            botao_ok = self.esperar_elemento('/html/body/div[5]/div/div[2]/div[3]/input[2]')
            if botao_ok:
                botao_ok.click()
                time.sleep(TEMPOS['wait'])  # Aumentado para 2 segundos
                print("  ✓ Botão OK clicado")

            # Clicando no botão Fechar
            print("  → Clicando no botão Fechar...")
            botao_fechar = self.esperar_elemento('/html/body/div[5]/div/div[2]/div[3]/input')
            if botao_fechar:
                botao_fechar.click()
                time.sleep(TEMPOS['wait'])  # Aumentado para 2 segundos
                print("  ✓ Botão Fechar clicado")

            print("  ✓ Sequência de ações completada com sucesso")

            return True

        except Exception as e:
            print(f"\n✗ Erro ao acessar importação: {e}")
            return False

    def iniciar_automacao(self):
        """Inicia o processo de automação."""
        try:
            print("\n=== Iniciando Configuração ===")
            print("→ Carregando credenciais...")
            self.credenciais = GerenciadorCredenciais.carregar_credenciais("login_imap.txt")
            if not self.credenciais:
                raise Exception("Falha ao carregar credenciais")
            print("✓ Credenciais carregadas com sucesso")

            print("\n→ Iniciando navegador...")
            self.navegador = ConfiguracaoNavegador.iniciar_navegador()
            if not self.navegador:
                raise Exception("Falha ao iniciar navegador")
            print("✓ Navegador iniciado com sucesso")

            if self.fazer_login():
                print("\n✓ Login realizado com sucesso!")
                
                print("\n→ Aguardando carregamento da página...")
                time.sleep(TEMPOS['tempo'])
                
                if self.acessar_importacao_usuarios():
                    print("\n✓ Processo de importação configurado com sucesso!")
                else:
                    print("\n✗ Falha na configuração da importação")
            else:
                print("\n✗ Falha no processo de login")

        except Exception as e:
            print(f"\n✗ Erro durante a automação: {e}")
        finally:
            if self.navegador:
                print("\n=== Finalizando Automação ===")
                input("\nPressione Enter para fechar o navegador...")
                self.navegador.quit()
                print("✓ Navegador fechado")

def main():
    """Função principal."""
    # Configurações iniciais
    os.environ['TF_CPP_MIN_LOG_LEVEL'] = '2'
    sys.stdout.reconfigure(encoding='utf-8')
    
    try:
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    except locale.Error:
        pass

    # Inicia automação
    automacao = AutomacaoPortal()
    automacao.iniciar_automacao()

if __name__ == "__main__":
    main() 