import pyautogui
import time 

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

NOME_ARQUIVO_CONFIG = "login_imap.txt"
caminho_pasta = obter_caminho_pasta(NOME_ARQUIVO_CONFIG)
caminho_pasta_downloads = obter_caminho_pasta(NOME_ARQUIVO_CONFIG)

delay = 0.7
tempo = 2
fast = 0.3

pyautogui.PAUSE = 0.2  # Pausa padrão entre comandos
pyautogui.FAILSAFE = True  # Move o mouse para o canto superior esquerdo para interromper


def move_and_click(x, y, duration=0.1):
    """Move o mouse para uma posição e clica."""
    pyautogui.moveTo(x, y, duration=duration)
    pyautogui.click()

def double_click(x, y, duration=0.1):
    """Move o mouse para uma posição e dá um duplo clique."""
    pyautogui.moveTo(x, y, duration=duration)
    pyautogui.doubleClick()

def press_keys(*keys, duration=0.1):
    """Pressiona uma sequência de teclas."""
    for key in keys:
        pyautogui.press(key)
        time.sleep(duration)

# -------------------------- Teste --------------------------
 

pyautogui.moveTo(827, 352)











