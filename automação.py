import pyautogui
import openpyxl
import keyboard
from time import sleep

pyautogui.PAUSE = 0.5
"""importar a planilha"""
planilha = openpyxl.load_workbook("planilha_email.xlsx")
planilha = planilha['Planilha1']

"""Abrir o navegador"""
pyautogui.press('win')
sleep(1)
pyautogui.write('Microsoft Edge')
sleep(1)
pyautogui.press('enter')
sleep(2)

"""abrir o email"""
pyautogui.write('gmail')
pyautogui.press('enter')
sleep(5)

"""escrever"""
posição = pyautogui.locateOnScreen('escrever.PNG')
sleep(2)
pyautogui.click(posição)

for linha in planilha.iter_rows(min_row=2):
    """guardar variaveis"""
    nome = linha[0].value
    email = linha[1].value
    assunto = linha[2].value
    mensagem = linha[3].value

    """email"""
    pyautogui.write(email)
    pyautogui.press('enter')
    sleep(0.5)
    pyautogui.press('tab')

    """assunto"""
    pyautogui.write(assunto)
    pyautogui.press('tab')
    sleep(0.5)

    """mensagem"""
    mensagem2 = f'Olá {nome}, {mensagem}'
    keyboard.write(mensagem2)
    sleep(0.5)

    """enviar"""
    enviar = pyautogui.locateOnScreen('enviar.PNG')
    sleep(2)
    pyautogui.click(enviar)

    """escrever dnv"""
    posição = pyautogui.locateOnScreen('escrever.PNG')
    sleep(2)
    pyautogui.click(posição)
