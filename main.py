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
        self.ids.grid.clear_widgets()  # Limpa os widgets existentes
        #abre arquivo das turmas cadastradas e pega de lá o nome das turmas
        try:
            #pega cada turma cadastrada
            arq_slvTurma = openpyxl.load_workbook('cadastro_turmas.xlsx')
            celulas = arq_slvTurma.active
            turmas_cadastradas = [celula.value for celula in celulas['A'][1:]]
        except FileNotFoundError:
            turmas_cadastradas = []
        
        #pega a quantidade de turmas
        num_turmas = len(turmas_cadastradas)
        dias_semana = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta']

        # pega a quantidade de turmas e cria os botões de selecao
        for turma in turmas_cadastradas:
            self.ids.grid.add_widget(Label(text=turma))
            for dia in dias_semana:
                spinner = Spinner(
                    text='Escolha',
                    values=[str(i) for i in range(1, num_turmas + 1)], 
                    size_hint=(None, None),
                    size=(150, 40),
                    pos_hint={'center_x': 0.5}
                )
                self.ids.grid.add_widget(spinner) # adiciona turma com os dias da semana na tela

    def salvar_Ordem(self):
            
        arq_slvDataOrdem = openpyxl.Workbook()
        celula = arq_slvDataOrdem.active
        celula.append(['Turma', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta'])

        divisao_grids = self.ids.grid.children
        num_turmas = len(divisao_grids) // 6
        
        # cria uma lista com as turmas e as escolhas
        for i in range(num_turmas):
            turma_index = i * 6 + 5 
            turma = divisao_grids[turma_index].text #pega o nome das turmas pela posicao
            escolhas = [divisao_grids[turma_index - j].text for j in range(1, 6)] #pega as escolhas pela posicao
            celula.append([turma] + escolhas) #adiciona na planilha
        
        arq_slvDataOrdem.save('cadastro_OrdemTurmas.xlsx')

        # Resetar os textos dos Spinners
        for i in range(len(divisao_grids)):
            if isinstance(divisao_grids[i], Spinner):
                divisao_grids[i].text = 'Escolha'

class RegistroDia(Screen):
    def on_pre_enter(self):
        tab_Monitor = openpyxl.load_workbook('cadastro_monitores.xlsx')
        celulas = tab_Monitor.active
        monitores_cadastrados = [celula.value for celula in celulas['A'][1:]]
        self.ids.spMonitor.values = monitores_cadastrados


class Relatorio(Screen):
    pass

class Mofome(App):
    def build(self):
        return Gerenciador()

if __name__ == '__main__':
    Mofome().run()