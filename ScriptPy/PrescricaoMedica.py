# IMPORTA AS LIB'S PARA O CODIGO FUNCIONAR
import subprocess
import sys
import win32com.client
import keyboard
import pyautogui
import time
import tkinter as tk

# DEFINE AS CONSTANTES DO SCRIPT
SAP_GUI_EXE_PATH = "C:\\Program Files (x86)\\SAP\\FrontEnd\\SapGui\\saplogon.exe" # DEFINE AONDE ESTÁ O SAP
SAP_ID = "/nztsd058"  # TRANSAÇÃO QUE DESEJA ACESSAR NO SAP
USER_ID = None  # USUÁRIO DO SAP
PASSWORD = None  # SENHA DO USUÁRIO
DATE_FROM = None  # CAMPO DATA (DE) - PREENCHIDA PELO USUÁRIO
DATE_TO = None  # CAMPO DATA (ATE) - PREENCHIDA PELO USUÁRIO

# PARTE DO CÓDIGO QUE SOLICITA O USUÁRIO SAP E SENHA QUE SERA ARMAZENADA NAS CONSTANTES "USER_ID e PASSWORD"
def save_credentials():
    global USER_ID, PASSWORD
    USER_ID = sap_entry.get()
    PASSWORD = password_entry.get()
    root.destroy()

def main():
    # TRRANSFORMA AS CONSTANTES EM NULAS
    global root, sap_entry, password_entry, USER_ID, PASSWORD
    USER_ID = None
    PASSWORD = None

    root = tk.Tk()
    sap_label = tk.Label(root, text="Insira o usuário SAP:")
    sap_label.pack(pady=5)
    sap_entry = tk.Entry(root)
    sap_entry.pack(pady=5)

    password_label = tk.Label(root, text="Insira a senha do SAP:")
    password_label.pack(pady=5)
    password_entry = tk.Entry(root, show="*")# DEIXA A SENHA COM OS ASTERISTICOS
    password_entry.pack(pady=5)

    save_button = tk.Button(root, text="Salvar", command=save_credentials)
    save_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()

# PARTE DO CÓDIGO AONDE ELE SOLICITA AS DATAS PARA O USUÁRIO
def get_date():
    # DEIXA AS CONSTANTES "DATE_FROM" & "DATE_TO" GLOBAIS PARA SEREM USADAS EM OUTRAS PARTES DO CODIGO
    global DATE_FROM, DATE_TO
    if DATE_FROM is None:  # VERIFICA SE AS DUAS CONSTANTES ESTÃO DECLARADAS COMO NULAS PARA RECEBERAM O VALOR INSERIDO PELO USUÁRIO
        # SE A CONSTANTE FOR NULA, "GET()" RECEBE O OBJETO "ENTRADA" PARA ASSIM RECEBER A DATA QUE O USUÁRIO INSERIU
        DATE_FROM = entry.get()
        # APAGA O CONTEÚDO DA LINHA ENTRADA PARA O USUÁRIO INSERIR UMA NOVA DATA
        entry.delete(0, tk.END)
        # SOLICITA PARA O USUÁRIO INSERIR A SEGUNDA DATA (ATE)
        label.config(text="Segunda Data (ATE) (DD.MM.AAAA)")
    else:
        # SE A CONSTANTE FOR NULA, ELA VAI RECEBER O VALOR QUE O USUÁRIO PASSOU PARA ELA
        DATE_TO = entry.get()
        window.destroy()  # FECHA A JANELA APÓS PREENCHIMENTO DAS DATAS

window = tk.Tk()  # INICIALIZA A JANELA PARA POR AS DATAS

label = tk.Label(window, text="Primeira Data (DE) (DD.MM.AAAA)") # SOLICITA AO USUÁRIO A PRIMEIRA DATA (DE)
label.pack()

entry = tk.Entry(window)
entry.pack()

# BOTÃO QUE RECEBE A FUNÇÃO DE CLICAR PARA INSERIR OS DADOS
button = tk.Button(window, text="Inserir", command=get_date)
button.pack()

window.mainloop()  # MANTÉM A JANELA EM LOOPING PARA NÃO FECHAR

winium = subprocess.Popen(SAP_GUI_EXE_PATH)# CAMINHO QUE ESTA LOCALIZADO O EXE DO SAP
time.sleep(1)

# CLICANDO NO AMBIENTE QUE DESEJA ACESSAR NO SAP GUI
class SapGui(object):
    def __init__(self):

        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        # SCRIPT DO PROPRIO SAP PARA EXECUTAR OUTRAS FUNÇÕES - EX: findById("wnd[0]").maximize
        application = self.SapGuiAuto.GetScriptingEngine

# ABRE CONEXÃO COM O SAP
        time.sleep(1)
        # CONEXÃO COM A ÁREA DE TRABALHO 'SAP PRODUÇÃO'
        self.connection = application.OpenConnection("SAP PRODUÇÃO", True)
        time.sleep(1)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize  # MAXIMIZA A TELA DO SAP

# LOGIN SENHA E USUARIO SAP
    def sapLogin(self):
        try:
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = USER_ID  # USER SAP
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = PASSWORD  # SENHA SAP
            pyautogui.hotkey('enter')  # CLICA NA TECLA ENTER
        except:
            print(sys.exc_info()[0])# MOSTRA SE ACONTECER ALGUM ERRO DURANTE A EXECUÇÃO

        time.sleep(1)
        self.register_supplier()# CHAMA A FUNÇÃO DE POR A TRANSAÇÃO DO SAP/ DATA/ COLAR PRODUTOS DE PRESCRIÇÃO

    # FUNÇÃO QUE ENDEREÇA A TRANSAÇÃO E PREENCHE A DATA NO SAP
    def register_supplier(self):

        # NÚMERO DA TRANSAÇÃO SAP
        self.session.findById("wnd[0]/tbar[0]/okcd").text = SAP_ID
        time.sleep(1)
        self.session.findById("wnd[0]").sendVKey(0)  # CLICA A TECLA ENTER
        self.session.findById(
            "wnd[0]/usr/btn%_P_MATNR_%_APP_%-VALU_PUSH").press

        # DATA NO CAMPO 'PERIÓDO' (DE)
        self.session.findById("wnd[0]/usr/ctxtP_EMIS-LOW").text = DATE_FROM
        # DATA NO CAMPO 'PERIÓDO' (ATE)
        self.session.findById("wnd[0]/usr/ctxtP_EMIS-HIGH").text = DATE_TO
        time.sleep(2)

        # VAI ATE A JANELA DE COLAR OS PRODUTOS DE PRESCRIÇÃO (ATUALMENTE SÃO 34 ITENS)
        for i in range(5):
            pyautogui.press('tab')
        pyautogui.hotkey('enter')
        # PRESSIONA A TECLA SHIFT+F12 PARA COLAR OS PRODUTOS
        self.session.findById("wnd[0]").sendVKey(24)
        time.sleep(1)
        self.session.findById("wnd[0]").sendVKey(8)  # PRESSIONA F8 PARA PROSSEGUIR
        time.sleep(1)

        # VAI ATE O CAMPO DE CHECK SELECIONANDO "FORMATO CLOSE-UP"
        for i in range(6):
            pyautogui.press('tab')
        pyautogui.hotkey("down")

        time.sleep(1)
        # PRESSIONA F8 PARA GERAR O RELÁTORIO :)
        self.session.findById("wnd[0]").sendVKey(8)

        # CHAMA A FUNÇÃO DE FECHAR O SAP
        self.close_window()

    # FUNÇÃO DE FECHAR O SAP APOS GERAR O RELATÓRIO
    def close_window(self):
        time.sleep(1)
        # PRESSIONA ENTER PARA FECHAR A TELA "ARQUIVOS GERADOS"
        pyautogui.hotkey('enter')

        # APERTA ALT+F4 PARA SAIR DO SAP
        while True:
            keyboard.press('ALT+F4')
            keyboard.release('ALT+F4')
            time.sleep(1)
            break

        pyautogui.hotkey('tab') # APERTA TAB PARA SELECIONAR A OPÇÃO 'sim' NO MENSSAGE BOX
        time.sleep(1)
        pyautogui.hotkey('enter') # APERTA ENTER NO 'sim' E SAI DO SAP

# CHAMA A FUNÇÃO PARA ABRIR O SAP, FAZER TODA A AUTOMATIZAÇÃO E REALIZA O LOGIN
SapGui().sapLogin()