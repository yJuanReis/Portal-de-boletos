import pyautogui
import time 

delay = 0.7
tempo = 2
fast = 0.3


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


# Defina o tempo de espera entre as ações (em segundos)
fast = 0.3


def press_key(key):
    pyautogui.press(key)
    time.sleep(fast)

pyautogui.moveTo(507, 252)
pyautogui.click()
time.sleep(fast)










































