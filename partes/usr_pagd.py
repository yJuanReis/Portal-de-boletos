


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
time.sleep(tempo) 
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

