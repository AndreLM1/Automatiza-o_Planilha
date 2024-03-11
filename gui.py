
from __future__ import print_function
import os.path

from tkinter import *
import tkinter 

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Criacao da interface
interface = tkinter.Tk()

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Classe para as acoes 
class Action():

    # Funcao para adicionar acao aos botoes
    def limpa_tela(self):
        self.abertura_entry.delete(0, END)
        self.user_entry.delete(0, END)
        self.analista_entry.delete(0, END)
        self.setor_entry.delete(0, END)
        self.solicitacao_entry.delete(0, END)
        self.resolucao_entry.delete(0, END)
        self.fechamento_entry.delete(0, END)
        self.tipo_entry.delete(0, END)
        self.status_entry.delete(0, END)

    def add_chamado(self):
        self.data_abertura = self.abertura_entry.get()
        self.user = self.user_entry.get()
        self.analista = self.analista_entry.get()
        self.setor = self.setor_entry.get()
        self.solicitacao = self.solicitacao_entry.get()
        self.resolucao = self.resolucao_entry.get()
        self.data_fechamento = self.fechamento_entry.get()
        self.tipo = self.tipo_entry.get()
        self.status = self.status_entry.get()
        # print(self.st)
        main(self.data_abertura, self.user, self.analista, self.setor, self.solicitacao, self.resolucao, self.data_fechamento, self.tipo, self.status)


# Classe de Aplicacao para a Interface
class Application(Action):
    def __init__(self):
        self.interface = interface
        self.tela()
        self.botoes()
        self.frame()
        # Manter a interface aberta 
        interface.mainloop()
    
    # Funcao das caracteristicas
    def tela(self):
        self.interface.title("Chamados Totvs")
        self.interface.configure(background="#bbc4be")
        # Modificacao do tamanho
        interface.geometry("500x400")
        # Metodo para mudar o tamanho da tela
        self.interface.resizable(False, False)
    
    # Funcao dos Botoes 
    def botoes(self):
        # Botao limpar
        self.bt_limpar = Button(text= "Limpar Campos", command=self.limpa_tela)
        # Coordenadas do botao x e y
        self.bt_limpar.place(relx= 0.2, rely= 0.93, relwidth=0.2 )
        # Botao inserir
        self.bt_inserir = Button(text= "Inserir", command=self.add_chamado)
        # Coordenadas do botao x e y
        self.bt_inserir.place(relx= 0.6, rely= 0.93, relwidth=0.2 )
    
    # Funcao das Labels & Campos de preenchimento
    def frame(self):
        # Criacao de Labels
        # Data de Abertura
        self.data_abertura = Label(text= "Data de Abertura")
        self.data_abertura.place(relx=0.15,rely=0.05,relwidth=0.3)

        # Usuario
        self.user = Label(text= "Usuário")
        self.user.place(relx=0.15,rely=0.15,relwidth=0.3)

        # Analista
        self.analista = Label(text= "Analista")
        self.analista.place(relx=0.15,rely=0.25,relwidth=0.3)

        # Setor
        self.setor = Label(text= "Setor")
        self.setor.place(relx=0.15,rely=0.35,relwidth=0.3)

        # Solicitação
        self.solicitacao = Label(text= "Solicitação")
        self.solicitacao.place(relx=0.15,rely=0.45,relwidth=0.3)

        # Resolução
        self.resolucao = Label(text= "Resolução")
        self.resolucao.place(relx=0.15,rely=0.55,relwidth=0.3)

        # Data Fechamento
        self.data_fechamento = Label(text= "Data de Fechamento")
        self.data_fechamento.place(relx=0.15,rely=0.65,relwidth=0.3)

        # Tipo
        self.tipo = Label(text= "Tipo")
        self.tipo.place(relx=0.15,rely=0.75,relwidth=0.3)

        # Status 
        self.status = Label(text= "Status")
        self.status.place(relx=0.15,rely=0.85,relwidth=0.3)


        # Criacao de campos de texto (entry)
        # Data de Abertura
        self.abertura_entry = Entry()
        self.abertura_entry.place(relx=0.5,rely=0.05,relwidth=0.3)

        # Usuario
        self.user_entry = Entry()
        self.user_entry.place(relx=0.5,rely=0.15,relwidth=0.3)

        # Analista
        self.analista_entry = Entry()
        self.analista_entry.place(relx=0.5,rely=0.25,relwidth=0.3)

        # Setor
        self.setor_entry = Entry()
        self.setor_entry.place(relx=0.5,rely=0.35,relwidth=0.3)

        # Solicitação
        self.solicitacao_entry = Entry()
        self.solicitacao_entry.place(relx=0.5,rely=0.45,relwidth=0.3)

        # Resolução
        self.resolucao_entry = Entry()
        self.resolucao_entry.place(relx=0.5,rely=0.55,relwidth=0.3)

        # Data Fechamento
        self.fechamento_entry = Entry()
        self.fechamento_entry.place(relx=0.5,rely=0.65,relwidth=0.3)

        # Tipo
        self.tipo_entry = Entry()
        self.tipo_entry.place(relx=0.5,rely=0.75,relwidth=0.3)

        # Status
        self.status_entry = Entry()
        self.status_entry.place(relx=0.5,rely=0.85,relwidth=0.3)

# Funcao para ver qual a proxima linha na planilha vazia
def get_next_available_row(sheet):
    # acessa e percorre a planilha
    result = sheet.values().get(spreadsheetId="1EU7FZ9DB9J_4U81PKBWUnykywiyEJJ0svYL1frGv9PY", range="Chamados!A:I").execute()
    values = result.get('values', []) # pega os dados da linha atual 
    if not values:
        return 1  # Se não houver dados, a primeira linha disponível é a 1
    else:
        return len(values)+ 1   # Adiciona 1 para a próxima linha disponível


def main(data_abertura, user, analista, setor, solicitacao, resolucao, data_fechamento, tipo, status):
  # variavel creds comeca como vazia (que nao há ningúem logado)
  creds = None
 
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
      

    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "client_secret.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())
 
  service = build('sheets', 'v4', credentials=creds)
  # # Call the Sheets API
  # # Ler informações do Google Sheets
  sheet = service.spreadsheets()

  linha_atual = get_next_available_row(sheet)
  # A planilha não pode ser arquivo .xls, tem q salvar como planilha mesmo
  # File -> Salvar como Planilhas Google
  result = sheet.values().get(spreadsheetId="1EU7FZ9DB9J_4U81PKBWUnykywiyEJJ0svYL1frGv9PY",range="Chamados!A2:I").execute()
  values = result.get('values', [])
  
  # Adicionando novos chamados
  novo_chamado = [
     
    [ data_abertura, user, analista, setor, solicitacao, resolucao, data_fechamento, tipo, status]
  ]
  range_atual = f"Chamados!A{linha_atual}:I{linha_atual}"  # Constrói o intervalo com base na linha_atual
  result = sheet.values().update(spreadsheetId="1EU7FZ9DB9J_4U81PKBWUnykywiyEJJ0svYL1frGv9PY",range=range_atual,valueInputOption="USER_ENTERED",
                                  body={"values": novo_chamado}).execute()
  

Application()