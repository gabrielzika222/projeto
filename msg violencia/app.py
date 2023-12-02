import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 

webbrowser.open('https://web.whatsapp.com/')
sleep(15)

workbook = openpyxl.load_workbook('usuarios.xlsx')
pagina_usuarios = workbook['Sheet1']

for linha in pagina_usuarios.iter_rows(min_row=2):

    nome = linha[0].value
    cpf = linha[1].value
    endereco = linha[2].value
    telefone_emergencia = linha[3].value
    
    mensagem = f'Olá meu nome é {nome} meu CPF é {cpf} *e estou sofrendo com violencia domestica nesse exato momento, meu endereço é {endereco} e necessito de ajuda com urgencia*'


    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone=5548988440011&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_whatsapp)
    sleep(15)
    pyautogui.press('enter')
    sleep(10)
    pyautogui.hotkey('ctrl','w')
    sleep(10)
    
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone=5561996100180&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_whatsapp)
    sleep(15)
    pyautogui.press('enter')
    sleep(10)
    pyautogui.hotkey('ctrl','w')
    sleep(10)

    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone_emergencia}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_whatsapp)
    sleep(15)
    pyautogui.press('enter')
    sleep(10)
    pyautogui.hotkey('ctrl','w')
    