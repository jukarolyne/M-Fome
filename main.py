import openpyxl
from kivy.app import App
from kivy.uix.screenmanager import Screen, ScreenManager
from kivy.lang import Builder
from kivy.uix.label import Label
from kivy.uix.spinner import Spinner

class Gerenciador(ScreenManager):
    pass

class Menu(Screen):
    pass

class Cadastro(Screen):
    pass

class CadastroTurmas(Screen):
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

        self.ids.nomeTurma.text = ''
        self.ids.matriculaTurma.text = ''

class CadastroMonitor(Screen): 
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

        self.ids.nomeAluno.text = ''  

class CadastroOrdem(Screen):
    def on_pre_enter(self):
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
            self.ids.grid.add_widget(Label(text=dia, 
                                           size_hint_y=None,
                                           height= 40))
            

        for turma in turmas_cadastradas:
            self.ids.grid.add_widget(Label(text=turma,
                                           size_hint_y=None,
                                           height=40))
            for _ in dias_semana[1:]:
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
            escolhas = [divisao_grids[turma_index - j].text for j in range(1, 6)]
            celula.append([turma] + escolhas)
        
        arq_slvDataOrdem.save('cadastro_OrdemTurmas.xlsx')

        for i in range(len(divisao_grids)):
            if isinstance(divisao_grids[i], Spinner):
                divisao_grids[i].text = 'Escolha'

class RegistroDia(Screen):
    def on_pre_enter(self):
        tab_Monitor = openpyxl.load_workbook('cadastro_monitores.xlsx')
        celulas = tab_Monitor.active
        monitores_cadastrados = [celula.value for celula in celulas['A'][1:]]
        self.ids.spMonitor.values = monitores_cadastrados
    
    def salvar_frequencia(self, data, almoco, monitor):
        try:
            workbook = openpyxl.load_workbook('frequencia.xlsx')
        except FileNotFoundError:
            workbook = openpyxl.Workbook()
            celula = workbook.active
            celula.title = 'Frequencia'
            celula.append(['Data', 'Almoço', 'Monitor'])
        else:
            celula = workbook['Frequencia']

        celula.append([data, almoco, monitor])

        workbook.save('frequencia.xlsx')
        #print("Frequência salva com sucesso.")

        self.ids.data.text = ''
        self.ids.almoco.text = ''
        self.ids.spMonitor.text = 'Escolha Monitor'

class Relatorio(Screen):
    pass

class Mofome(App):
    def build(self):
        return Gerenciador()

if __name__ == '__main__':
    Mofome().run()