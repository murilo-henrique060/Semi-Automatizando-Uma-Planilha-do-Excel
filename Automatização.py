from datetime import date
from openpyxl import load_workbook
import PySimpleGUI as sg
from os import getcwd
from time import localtime

base_path = getcwd()

def convert_number_alpha(intValue):
    intValue = int(intValue)
    Value = intValue % 26
    letras = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    letra = letras[Value]
    return letra

class Planilha():
    def __init__(self):
        self.path = base_path + '\\Rendimento Mensal Familiar Automatizado.xlsx'

        self.planilha = load_workbook(self.path)
        self.Gastos_Anuais = self.planilha.active
        self.Projecao_Planilha = []
        self.contas = []
        self.data = []
        self.linha_Planilha = 0
        self.valores_fixos = []
        self.cordenadas = []

        for rows in self.Gastos_Anuais.rows:
            self.Projecao_Planilha.append([])
            for cell in rows:
                self.Projecao_Planilha[self.linha_Planilha].append(cell.value if cell.value != None else ' ')
            self.linha_Planilha += 1

        self.linha_Planilha = 0
        self.planilha_size = len(self.Projecao_Planilha)

        for linhas in self.Projecao_Planilha:
            if self.linha_Planilha > 1 and self.linha_Planilha < self.planilha_size - 1:
                self.contas.append(linhas[0])
                self.valores_fixos.append(['Selecionar'])
            self.linha_Planilha += 1

        self.linha_Planilha = 0
        self.itens_planilha = 0

        for linhas in range(len(self.contas)):
            self.editList = self.Projecao_Planilha[linhas + 2]
            self.editList.pop(0)
            self.editList.pop(-1)
            self.editList.pop(-1)
            self.data.append(self.editList)

    def excelUpdateData(self):
        for linha in range(len(self.data)):
            for itenLinha in range(12):
                pos = convert_number_alpha(itenLinha + 1) + str(linha + 3)
                self.Gastos_Anuais[pos] = self.data[linha][itenLinha]

    def adicionar(self,data,contas,linhaConta,month):
        addValor = 0

        if data['Select_Adicionar_Valor_'+f'{contas}'] and str(data['Adicionar_valor_' + f'{contas}']).strip().replace('.','').replace(',','').isnumeric():
            try:
                addValor = float(str(data['Adicionar_valor_' + f'{contas}']).strip().replace(',','.'))
            
            except:
                addValor = 0

        elif data['Select_Adicionar_Valor_'+f'{contas}'] == False and str(data['Valor_fixo_'+f'{contas}']).replace('.','').replace(',','').isnumeric():
            try:
                addValor = float(str(data['Valor_fixo_'+f'{contas}']).strip().replace(',','.'))
            
            except:
                addValor = 0

        valorAtual = self.data[linhaConta][month - 1]

        if str(valorAtual).strip() == '':
            valorAtual = 0

        self.data[linhaConta][month - 1] = float(valorAtual) + addValor

        self.excelUpdateData()

planilha = Planilha()

class Window():
    def __init__(self):
        # Data
        self.icon_path = base_path + '\\xls-icon-3392-Windows.ico'

        self.year = localtime().tm_year
        self.month = localtime().tm_mon
        self.day = localtime().tm_mday

        self.date = ()

        # Layout
        sg.theme('DarkAmber')
        sg.set_global_icon(self.icon_path)

        self.layout = [
    [sg.Text(f'Rendimento Mensal Familiar Automatizado.xlsx'),sg.Text(f'{f"{planilha.Projecao_Planilha[1][self.month]} de {self.year}":>30}',(30,0),enable_events=True,key='Date_Input')],
    [sg.Text('-' * 200)]

        ]

        self.layout2 = []

        linhaConta = 0

        for contas in planilha.contas:
            self.layout.append([sg.Text(f'{contas}',(30,0)),sg.Text('-'),sg.Text(f'{planilha.data[linhaConta][self.month - 1]:>9}',(9,0),k='Valor_'+f'{contas}'),sg.Button('Adicionar Valor',k=f'{contas}')])
            linhaConta += 1

        # Window
        self.leigth = 500
        self.height = 600

        self.window = sg.Window('Automatização Excel',self.layout,size=(self.leigth,self.height),return_keyboard_events=True)

        self.window_2 = None

        self.event2 = None
        self.value2 = None

    def update(self):
        self.event, self.values = self.window.read()

    def update2(self,contas):
        if self.value2["Select_Adicionar_Valor_"+f"{contas}"] != None:
            if self.value2['Select_Adicionar_Valor_'+f'{contas}']:
                self.window_2['Adicionar_valor_' + f'{contas}'].update(disabled=False)
                self.window_2['Tornar_Fixo_'+f'{contas}'].update(disabled=False)
                self.window_2['Valor_fixo_'+f'{contas}'].update(disabled=True)
            
            else:
                self.window_2['Adicionar_valor_' + f'{contas}'].update(disabled=True)
                self.window_2['Tornar_Fixo_'+f'{contas}'].update(disabled=True)
                self.window_2['Valor_fixo_'+f'{contas}'].update(disabled=False)

    def update3(self,contas):
        self.window.Element('Valor_'+f'{contas}').Update(f'{planilha.data[linhaConta][screen.month - 1]:>9}')
        self.window_2.Element('Valor_Atual'+f'{contas}').Update(f'{f"{planilha.data[linhaConta][screen.month - 1]}":^115}')
        self.window_2.Element('Adicionar_valor_' + f'{contas}').Update('')

    def dateChange(self):
        self.date = sg.popup_get_date(title='Escolha uma Data')

        self.day = self.date[1]
        self.month = self.date[0]
        self.year = self.date[2]

screen = Window()

# Main Loop
while True:
    screen.update()
    if screen.event == sg.WIN_CLOSED:
        break

    if screen.event == 'Date_Input':
        screen.dateChange()
        screen.window.Element('Date_Input').Update(f'{f"{planilha.Projecao_Planilha[1][screen.month]} de {screen.year}":>30}')

        linhaConta = 0

        for contas in planilha.contas:
            screen.window.Element('Valor_'+f'{contas}').Update(f'{planilha.data[linhaConta][screen.month - 1]:>9}')
            linhaConta += 1

    linhaConta = 0


    for contas in planilha.contas:
        if screen.event == f'{contas}':
            screen.window.hide()
            
            screen.layout2 = [
                [sg.Text(f'Rendimento Mensal Familiar Automatizado.xlsx'),sg.Text(f'{f"{planilha.Projecao_Planilha[1][screen.month]} de {screen.year}":>30}',(30,0))],
                [sg.Text('-'*115)],
                [sg.Text(f'{contas:^110}',font=50)],
                [sg.Text('-'*115)],
                [sg.Text(f'{"Valor Atual":^115}')],
                [sg.Text(f'{f"{planilha.data[linhaConta][screen.month - 1]}":^115}',font=100,key='Valor_Atual'+f'{contas}')],
                [sg.Text('')],
                [sg.Text('Nome do Gasto (opcional): ',auto_size_text=True),sg.Input(size=(20,0),key='Nome_Gasto_'+f'{contas}')],
                [sg.Radio('Adicionar Valor:','Values',default=True,key='Select_Adicionar_Valor_'+f'{contas}'),sg.Text('   R$'),sg.Input(key='Adicionar_valor_' + f'{contas}',size=(10,0)),sg.Text('     '),sg.Checkbox('Tornar Fixo',key='Tornar_Fixo_'+f'{contas}')],
                [sg.Radio(f'{"Valor Fixo:":<18}','Values',key='Select_Valor_Fixo_'+f'{contas}'),sg.Text('   R$'),sg.Combo(planilha.valores_fixos[linhaConta],default_value='Selecionar',size=(10,0),key='Valor_fixo_'+f'{contas}',disabled=True)],
                [sg.Button('Adicionar',auto_size_button=True,key='add_'+f'{contas}'),sg.Text('  '),sg.Button('Voltar',auto_size_button=True,key='voltar_'+f'{contas}')]
            ]

            screen.window_2 = sg.Window(f'{contas}',screen.layout2,size=(screen.leigth,screen.height),return_keyboard_events=True)

            controle = True

            while True:
                screen.event2, screen.value2 = screen.window_2.read(50)

                if screen.event2 == sg.WIN_CLOSED or screen.event2 == 'Escape:27' or screen.event2 == 'voltar_'+f'{contas}':
                    controle = False
                    break 

                if screen.event2 == 'add_'+f'{contas}' or screen.event2.strip() == '' and screen.event2 != ' ':
                    planilha.adicionar(screen.value2,contas,linhaConta,screen.month)
                    screen.update3(contas)

                if controle:
                    screen.update2(contas)

            screen.window_2.close()
            screen.window.un_hide()

        linhaConta += 1

screen.window.close()

planilha.planilha.save(planilha.path)
planilha.planilha.close()
