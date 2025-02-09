import openpyxl
from kivy.app import App
from kivy.uix.screenmanager import Screen, ScreenManager
from kivy.lang import Builder
from kivy.uix.label import Label
from kivy.uix.spinner import Spinner
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
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

        except FileNotFoundError:
            turmas_cadastradas = []
        
        for turmas in turmas_cadastradas:
            self.ids.box.add_widget(Label(text=turmas, size_hint_y=None, size_hint_x=None, height=40))

    def salvar_dadosTurma(self, nome_turma, matricula_turma):
        try:
            tab_slvDataTurma = openpyxl.load_workbook('cadastro_turmas.xlsx')
        except FileNotFoundError:
            tab_slvDataTurma = openpyxl.Workbook()
            celula = tab_slvDataTurma.active
            celula.append(['Nome', 'Matrícula'])
        
        celula = tab_slvDataTurma.active
        celula.append([nome_turma, int(matricula_turma)])
        tab_slvDataTurma.save('cadastro_turmas.xlsx')

        self.ids.box.add_widget(Label(text=nome_turma, size_hint_y=None, height=40))

        self.ids.nomeTurma.text = ''
        self.ids.matriculaTurma.text = ''

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
            #btn_remover = Button(text='Excluir', size_hint_y=None, height=40)
            #self.ids.box.add_widget(btn_remover)

    def remover_monitor(self, monitor):  
        pass    

    def salvar_dadosMonitor(self, nome_aluno):

        try:
            arq_slvMonitor = openpyxl.load_workbook('cadastro_monitores.xlsx')
        except FileNotFoundError:
            arq_slvMonitor = openpyxl.Workbook()
            celula = arq_slvMonitor.active
            celula.append(['Nome'])
        
        celula = arq_slvMonitor.active
        celula.append([nome_aluno])
        arq_slvMonitor.save('cadastro_monitores.xlsx')

        self.ids.box.add_widget(Label(text=nome_aluno, size_hint_y=None, height=40))

        self.ids.nomeAluno.text = ''  

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

            for j in range(1, 6):
                escolhas.append(divisao_grids[turma_index - j].text)

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
        monitores_cadastrados = [celula.value for celula in celulas['A'][1:]]
        self.ids.spMonitor.values = monitores_cadastrados

        try:
            arq_slvTurma = openpyxl.load_workbook('cadastro_turmas.xlsx')
            celulas = arq_slvTurma.active
            turmas_cadastradas = [celula.value for celula in celulas['A'][1:]]
        except FileNotFoundError:
            turmas_cadastradas = []
        
        for turma in turmas_cadastradas:
            self.ids.grid.add_widget(Label(text=turma, 
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
            #celula.title = 'Frequencia'
            celula.append(['Data', 'Almoço', 'Monitor', 'Dia da Semana'])

        celula = arq_slvFrequencia.active
        celula.append([data, almoco, monitor, dia_semana])

        arq_slvFrequencia.save('frequencia.xlsx')

        self.ids.data.text = ''
        self.ids.almoco.text = ''
        self.ids.spMonitor.text = 'Escolha Monitor'
        self.ids.spDiaSemana.text = 'Escolha Dia'

class Relatorio(Screen):
    pass

class Mofome(App):
    def build(self):
        Window.size = (360, 640)
        return Gerenciador()

if __name__ == '__main__':
    Mofome().run()