import openpyxl
import matplotlib.pyplot as plt
from kivy.app import App
from kivy.uix.screenmanager import Screen, ScreenManager
from kivy.lang import Builder
from kivy.uix.label import Label
from kivy.uix.spinner import Spinner
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy_garden.matplotlib import FigureCanvasKivyAgg
from kivy.core.window import Window

class Gerenciador(ScreenManager):
    pass

class Menu(Screen):
    pass

class Cadastro(Screen):
    pass

class CadastroTurmas(Screen):
    def on_pre_enter(self):
        self.ids.box.clear_widgets()

        try:
            arq_slvTurmas = openpyxl.load_workbook('cadastro_turmas.xlsx')
            celulas = arq_slvTurmas.active
            turmas_cadastradas = [celula.value for celula in celulas['A'][1:]]
            qtde_alunos = [celula.value for celula in celulas['B'][1:]]

        except FileNotFoundError:
            turmas_cadastradas = []
            qtde_alunos = []
        
        for turma, alunos in zip(turmas_cadastradas, qtde_alunos):
            self.ids.box.add_widget(Label(text=f'TURMA: {turma} - QUANTIDADE DE ALUNOS: {alunos}', 
                                        size_hint_y=None,  
                                        height=40))
    
    def mostrar_popup(self, titulo, texto):
        msg_erro=Popup(
            title = titulo,
            content = Label(text=texto),
            size_hint=(None, None),
            size = (300, 200),
            padding=(10, 10, 10, 10)
        )

        msg_erro.open()

    def salvar_dadosTurma(self, nome_turma, qtde_alunos):
        try:
            tab_slvDataTurma = openpyxl.load_workbook('cadastro_turmas.xlsx')
        except FileNotFoundError:
            tab_slvDataTurma = openpyxl.Workbook()
            celula = tab_slvDataTurma.active
            celula.append(['Nome da Turma', 'Quantidade de Alunos'])
        
        celula = tab_slvDataTurma.active

        if len(nome_turma) > 0 and len(nome_turma) <= 5 and qtde_alunos.isdigit() and 1 <= int(qtde_alunos) <= 99:
            celula.append([nome_turma.upper(), int(qtde_alunos)])
            self.ids.box.add_widget(Label(text=f'TURMA: {nome_turma.upper()} - QUANTIDADE DE ALUNOS: {qtde_alunos}', 
                                      size_hint_y=None, 
                                      height=40))
            self.mostrar_popup('Sucesso!', "A Turma solicitada foi adicionada.")

        elif ValueError:
            if qtde_alunos == '' or nome_turma == '':
                self.mostrar_popup('Campo Vazio ou Valor Inválido!', 'Digite um valor válido. \nTurma deve no máximo 5 caracteres; \nQuantidade de Pessoas deve \nser de 1 a 99.')
            else:
                self.mostrar_popup('Campo Vazio ou Valor Inválido!', 'Digite um valor válido. \nTurma deve no máximo 5 caracteres; \nQuantidade de Pessoas deve \nser de 1 a 99.')

        tab_slvDataTurma.save('cadastro_turmas.xlsx')

        self.ids.nomeTurma.text = ''
        self.ids.qtdeAlunos.text = ''
    
    def remover_turma(self, nome_turma):
        try:
            tab_slvDataTurma = openpyxl.load_workbook('cadastro_turmas.xlsx')
        except FileNotFoundError:
            return self.mostrar_popup('Erro!', 'Não há turmas cadastradas. Cadastre uma turma primeiro!')
            
        celula = tab_slvDataTurma.active
        
        for row in celula.iter_rows(min_row=2, max_col=1):
            if row[0].value == nome_turma.upper():
                celula.delete_rows(row[0].row, 1)
                self.mostrar_popup('Sucesso!', 'A Turma solicitada foi removida.')
                break
        else:
            self.mostrar_popup('Campo Vazio ou Valor Inválido!', 'Essa turma não existe ou você \ndeixou o campo vazio. \nConfira as turmas cadastradas \nacima antes de remover.')

        tab_slvDataTurma.save('cadastro_turmas.xlsx')

        self.ids.nomeTurma.text = ''
        self.on_pre_enter()

class CadastroMonitor(Screen): 
    def on_pre_enter(self):
        self.ids.box.clear_widgets()

        try:
            arq_slvMonitor = openpyxl.load_workbook('cadastro_monitores.xlsx')
            celulas = arq_slvMonitor.active
            monitores_cadastrados = [celula.value for celula in celulas['A'][1:]]

        except FileNotFoundError:
            monitores_cadastrados = []
        
        for monitor in monitores_cadastrados:
            nome_monitor = Label(text=monitor, size_hint_y=None, height=40)
            self.ids.box.add_widget(nome_monitor)
    
    def mostrar_popup(self, titulo, texto):
        msg_erro=Popup(
            title = titulo,
            content = Label(text=texto),
            size_hint=(None, None),
            size = (300, 200),
            padding=(10, 10, 10, 10)
        )

        msg_erro.open()

    def salvar_monitor(self, nome_aluno):

        try:
            arq_slvMonitor = openpyxl.load_workbook('cadastro_monitores.xlsx')
        except FileNotFoundError:
            arq_slvMonitor = openpyxl.Workbook()
            celula = arq_slvMonitor.active
            celula.append(['Nome'])
        
        celula = arq_slvMonitor.active
        
        if len(nome_aluno)>= 3 and len(nome_aluno) <= 80:
            celula.append([nome_aluno.upper()])
            self.ids.box.add_widget(Label(text=nome_aluno.upper(), size_hint_y=None, height=40))

        else:
            self.mostrar_popup('Campo Vazio ou Valor Inválido!', 'Você digitou um nome inválido \nou deixou o campo vazio. \nDigite novamente!')
        
        arq_slvMonitor.save('cadastro_monitores.xlsx')

        self.ids.nomeAluno.text = ''  
    
    def remover_monitor(self, nome_aluno):  
        try:
            arq_slvMonitor = openpyxl.load_workbook('cadastro_monitores.xlsx')
        except FileNotFoundError:
            return 'Não há nomes cadastrados para serem excluídos!'
        
        celula = arq_slvMonitor.active
        for row in celula.iter_rows(min_row=2, max_col=1):
            if row[0].value == nome_aluno.upper():
                celula.delete_rows(row[0].row, 1)
                self.mostrar_popup('Sucesso!', 'O Monitor solicitado foi removido.')
                break
        else:
            self.mostrar_popup('Erro de Digitação!', 'Esse nome não existe. \nConfira os monitores cadastrados\nacima antes de remover.')

        arq_slvMonitor.save('cadastro_monitores.xlsx')

        
        self.ids.nomeAluno.text = ''  
        self.on_pre_enter()

class CadastroOrdem(Screen):
    def on_pre_enter(self):
        Window.size = (1240, 480)
        self.ids.grid.clear_widgets()
        try:
            arq_slvTurma = openpyxl.load_workbook('cadastro_turmas.xlsx')
            celulas = arq_slvTurma.active
            turmas_cadastradas = [celula.value for celula in celulas['A'][1:]]
        except FileNotFoundError:
            turmas_cadastradas = []
        
        num_turmas = len(turmas_cadastradas)
        dias_semana = ['Turma','Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta']

        for dia in dias_semana:
            self.ids.grid.add_widget(Label(
                text=dia, 
                size_hint_y=None,
                height= 40))

        for turma in turmas_cadastradas:
            self.ids.grid.add_widget(Label(
                text=turma,
                size_hint_y=None,
                height=40))
            for dia in dias_semana[1:]:
                    self.ids.grid.add_widget(Spinner(
                        text='Escolha', 
                        values=[str(i) for i in range(1, num_turmas + 1)],
                        size_hint=(None, None),
                        size=(150, 40),
                        pos_hint={'center_x': 0.5}))

    def salvar_Ordem(self):
        arq_slvDataOrdem = openpyxl.Workbook()
        celula = arq_slvDataOrdem.active
        celula.append(['Turma', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta'])

        divisao_grids = self.ids.grid.children
        num_turmas = len(divisao_grids) // 6
        
        for i in range(num_turmas):
            turma_index = i * 6 + 5 
            turma = divisao_grids[turma_index].text 
            escolhas = []

            if not turma == 'Turma':
                for j in range(1, 6):
                    valor = divisao_grids[turma_index - j].text
                    if valor.isdigit():
                        escolhas.append(int(valor))

            celula.append([turma] + escolhas)
        
        arq_slvDataOrdem.save('cadastro_OrdemTurmas.xlsx')

        for i in range(len(divisao_grids)):
            if isinstance(divisao_grids[i], Spinner):
                divisao_grids[i].text = 'Escolha'
    
    def on_leave(self):
        Window.size = (360, 640)

class RegistroDia(Screen):
    def on_pre_enter(self):
        tab_Monitor = openpyxl.load_workbook('cadastro_monitores.xlsx')
        celulas = tab_Monitor.active
        monitores_cadastrados = []
        for celula in celulas['A'][1:]:
            nome = celula.value.split(" ")
            nome1 = nome[0]
            nome2 = nome[-1]
            monitores_cadastrados.append(str(nome1+' '+nome2))

        self.ids.spMonitor.values = monitores_cadastrados
        self.ids.spDiaSemana.bind(text=self.on_spinner_select)
        self.ordenar_turmas_dia(self.ids.spDiaSemana.text)

    def on_spinner_select(self, spinner, text):
        self.ordenar_turmas_dia(text)
    
    def ordenar_turmas_dia(self, dia_semana):
        self.ids.grid.clear_widgets()
        try:
            arq_slvOrdem = openpyxl.load_workbook('cadastro_OrdemTurmas.xlsx')
            celulas = arq_slvOrdem.active
            dias_semana = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta']
            dia_semana_index = dias_semana.index(dia_semana) + 1  # +1 porque a primeira coluna é 'Turma'

            ordem_turmas = {}
            for celula in celulas.iter_rows(min_row=2):
                ordem_turmas[celula[0].value] = celula[dia_semana_index].value

        except FileNotFoundError:
            ordem_turmas = {}
        
        ordem_turmas = {turma: ordem for turma, ordem in ordem_turmas.items() if ordem is not None}
        
        turmas_ordenadas = sorted(ordem_turmas.items(), key=lambda x: x[1])
        
        for ordem, (turma, _) in enumerate(turmas_ordenadas, start=1):
            form_turma = f'{ordem}º - {turma}'
            self.ids.grid.add_widget(Label(text=form_turma, 
                                           size_hint_y=None, 
                                           height=40))
            self.ids.grid.add_widget(TextInput(hint_text='Meninos', 
                                               size_hint_y=None, 
                                               height=40,
                                               multiline=False))
            self.ids.grid.add_widget(TextInput(hint_text='Meninas', 
                                               size_hint_y=None, 
                                               height=40,
                                               multiline=False))
    
    def salvar_frequencia(self, data, almoco, monitor, dia_semana):
        try:
            arq_slvFrequencia = openpyxl.load_workbook('frequencia.xlsx')
        except FileNotFoundError:
            arq_slvFrequencia = openpyxl.Workbook()
            celula = arq_slvFrequencia.active
            celula.append(['Data', 'Almoço', 'Monitor', 'Dia da Semana', 'Turmas', 'Meninos', 'Meninas'])

        celula = arq_slvFrequencia.active

        divisao_grids = self.ids.grid.children
        num_turmas = len(divisao_grids) // 3
        
        for i in range(num_turmas):
            turma_index = i * 3 + 2 
            turma = divisao_grids[turma_index].text 
            escolhas = []

            for j in range(1, 3):
                escolhas.append(divisao_grids[turma_index - j].text)

            quantidade_meninos = divisao_grids[turma_index - 2].text
            quantidade_meninas = divisao_grids[turma_index - 1].text
            
            celula.append([data, almoco, monitor, dia_semana, turma, int(quantidade_meninos), int(quantidade_meninas)])

        arq_slvFrequencia.save('frequencia.xlsx')

        self.ids.data.text = ''
        self.ids.almoco.text = ''
        self.ids.spMonitor.text = 'Escolha Monitor'
        self.ids.spDiaSemana.text = 'Segunda'
        for i in range(len(divisao_grids)):
            if isinstance(divisao_grids[i], TextInput):
                divisao_grids[i].text = ''

class Relatorio(Screen):
    pass

class Mofome(App):
    def build(self):
        Window.size = (360, 640)
        return Gerenciador()

if __name__ == '__main__':
    Mofome().run()

