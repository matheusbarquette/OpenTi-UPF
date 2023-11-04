import cv2
import pyautogui
import numpy as np
import time 
import os 
import logging
import sys 

# Função para clicar em elementos
def click(img):
    try:
        # Carregar a imagem de tela completa e a imagem do template
        screenshot_full = np.array(pyautogui.screenshot())
        template = cv2.imread(img)

        # Converter screenshot para escala de cinza
        screenshot_gray = cv2.cvtColor(screenshot_full, cv2.COLOR_BGR2GRAY)

        # Realizar a correspondência de template
        result = cv2.matchTemplate(screenshot_gray, cv2.cvtColor(template, cv2.COLOR_BGR2GRAY), cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

        # Obter coordenadas do ponto correspondente
        x, y = max_loc

        # Calcular as coordenadas reais na tela
        template_height, template_width, _ = template.shape
        x_on_screen = x + template_width // 2
        y_on_screen = y + template_height // 2

        # Clicar na posição correspondente usando o PyAutoGUI
        pyautogui.click(x_on_screen, y_on_screen)
        time.sleep(2)

    except Exception as e:
        logging.error(f"Nao foi possivel clicar: {str(e)} ")
        print("Nao foi possivel clicar: " + str(e))

# Função para verificar se o elemento está visível
def elemento_visivel(element_img_path):

    cont = 0 

    while True:
        time.sleep(1)

        if cont == 3:
            print("nao foi possivel localizar elemento!")
            logging.error('nao foi possivel localizar elemento!')
            exit()
        
        screenshot = np.array(pyautogui.screenshot())
        element_template = cv2.imread(element_img_path)
        
        # Converter a imagem da tela para escala de cinza
        screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
        
        # Realizar a correspondência de template
        result = cv2.matchTemplate(screenshot_gray, cv2.cvtColor(element_template, cv2.COLOR_BGR2GRAY), cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        
        # Defina um limite para a correspondência (ajuste conforme necessário)
        limite = 0.9

        cont += 1

        # Verifique se o valor de correspondência é maior que o limite
        if max_val >= limite:
            print('visivel')
            return True
        else:
            print('nao visivel')

# configura o log do rpa 
logging.basicConfig(filename='rpa_log.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.DEBUG)
logging.getLogger().addHandler(console_handler)

logging.info('Iniciando RPA.')

# dir Base Images
PATH_EXE = os.getcwd()
DIR_IMAGES = os.path.join(PATH_EXE, 'src\images')

# clica no menu do windows
logging.info('clica no menu do windows')
click(fr"{DIR_IMAGES}\menu_iniciar.png")


# escreve notepad no executar
logging.info('escreve notepad no executar')
pyautogui.write("notepad")
time.sleep(1.5)

# Pressione a tecla enter
logging.info('Pressione a tecla enter')
pyautogui.hotkey('enter')
time.sleep(3.0)

# Texto a ser digitado
texto = 'Ola eu sou um robo participando do evento OPENTI da universidade UPF (:'

# Verifica se o Notepad esteja ativo
logging.info('Verifica se o Notepad esteja ativo')
elemento_visivel(fr"{DIR_IMAGES}\notepad.png")

# Digitar o texto caractere por caractere com pausas
for char in texto:
    pyautogui.write(char)
    time.sleep(0.03)  # Ajuste o tempo de pausa conforme desejadonotepad

# aguarda um tempo
logging.warning('aguarda um tempo')
time.sleep(1)

# pressiona CRTL + S para salvar o arquivo
logging.info('pressiona CRTL + S para salvar o arquivo')
pyautogui.hotkey('ctrl', 's')
time.sleep(3.0)

# Escolher diretorio para salvar o arquivo
logging.info(fr'Escolher diretorio para salvar o arquivo: C:\Users\{os.getlogin()}\Downloads\ola.txt')
pyautogui.write(fr'C:\Users\{os.getlogin()}\Downloads\ola.txt')
time.sleep(3.0)

# Pressione a tecla enter
logging.info('Pressione a tecla enter')
pyautogui.hotkey('enter')
time.sleep(3.0)

# Pressiona ALT + F4 para fechar o notepad
logging.info('Pressiona ALT + F4 para fechar o notepad')
pyautogui.hotkey('alt', 'f4')
time.sleep(3.0)

logging.info('Finalizado automacao!')
