import pyautogui
import pyautogui as escolha_opcao

opcao = pyautogui.confirm('Clique no botão desejado', buttons = ['Excel', 'Word', 'Google Chrome'])

if opcao == 'Excel':
    
    escolha_opcao.hotkey('ctrl', 'alt', 't')

    escolha_opcao.sleep(2)

    escolha_opcao.typewrite('libreoffice')

    escolha_opcao.sleep(6)

    escolha_opcao.press('Enter')

    escolha_opcao.sleep(2)

    escolha_opcao.click(x=162, y=412)

    escolha_opcao.sleep(2)

    escolha_opcao.typewrite('Escolhi a opção para abrir o excel')

elif opcao == 'Word':
    escolha_opcao.hotkey('ctrl', 'alt', 't')

    escolha_opcao.sleep(2)

    escolha_opcao.typewrite('libreoffice')

    escolha_opcao.sleep(6)

    escolha_opcao.press('Enter')

    escolha_opcao.sleep(2)

    escolha_opcao.click(x=221, y=349)

    escolha_opcao.sleep(2)

    escolha_opcao.typewrite('Escolhi a opção para abrir o Word')

elif opcao == 'Google Chrome':

    escolha_opcao.hotkey('ctrl', 'alt', 't')

    escolha_opcao.sleep(2)

    escolha_opcao.typewrite('google-chrome')

    escolha_opcao.sleep(2)

    escolha_opcao.press('Enter')

    escolha_opcao.sleep(3)

    escolha_opcao.typewrite('https://www.google.com/')

    escolha_opcao.sleep(2)

    escolha_opcao.press('Enter')