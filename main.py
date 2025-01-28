import openpyxl
from kivy.app import App
from kivy.uix.screenmanager import Screen, ScreenManager
from kivy.lang import Builder

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
            celula.append(['Nome', 'Turma'])
        
        celula = tab_slvDataTurma.active
        celula.append([nome_turma, int(matricula_turma)])
        tab_slvDataTurma.save('cadastro_turmas.xlsx')

        self.ids.nomeTurma.text = ''
        self.ids.matriculaTurma.text = ''

class CadastroAlunos(Screen): 
    def on_pre_enter(self):
        tab_Turmas = openpyxl.load_workbook('cadastro_turmas.xlsx')
        celulas = tab_Turmas.active
        turmas_cadastradas = [celula.value for celula in celulas['A'][1:]]
        self.ids.spTurma.values = turmas_cadastradas

    def salvar_dadosAluno(self, nome_aluno, turma_aluno, matricula_aluno, monitor, sexo):
        
        if monitor == False:
            monitor = 'Não'
        else:
            monitor = 'Sim'
        
        if sexo == 'Masculino':
            sexo = 'M'
        else:
            sexo = 'F'

        try:
            arq_slvAluno = openpyxl.load_workbook('cadastro_alunos.xlsx')
        except FileNotFoundError:
            arq_slvAluno = openpyxl.Workbook()
            celula = arq_slvAluno.active
            celula.append(['Nome', 'Turma', 'Matricula', 'Participa da Monitoria', 'Sexo'])
        
        celula = arq_slvAluno.active
        celula.append([nome_aluno, turma_aluno, int(matricula_aluno), monitor, sexo])
        arq_slvAluno.save('cadastro_alunos.xlsx')

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
        
        arq_slvDataOrdem.save('cadastro_OrdemTurmas.xlsx')

        self.ids.spSegunda.text = 'Escolha Turma'
        self.ids.spTerca.text = 'Escolha Turma'
        self.ids.spQuarta.text = 'Escolha Turma'
        self.ids.spQuinta.text = 'Escolha Turma'
        self.ids.spSexta.text = 'Escolha Turma'

class RegistroDia(Screen):
    pass

class Relatorio(Screen):
    pass

class Mofome(App):
    def build(self):
        return Gerenciador()

Builder.load_file('mofome.kv')
Mofome().run()