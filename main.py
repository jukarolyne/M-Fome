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
            arq_slvDataTurma = openpyxl.load_workbook('cadastro_turmas.xlsx')
        except FileNotFoundError:
            arq_slvDataTurma = openpyxl.Workbook()
            celula = arq_slvDataTurma.active
            celula.append(['Nome', 'Turma', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta'])
        
        celula = arq_slvDataTurma.active
        celula.append([nome_turma, int(matricula_turma)])
        arq_slvDataTurma.save('cadastro_turmas.xlsx')

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
        self.ids.matriculaAluno.text = ''
        self.ids.spTurma.text = 'Escolha Turma'
        self.ids.chkSim.active = False
        self.ids.chkNao.active = False
        self.ids.chkMasc.active = False
        self.ids.chkFem.active = False        

class CadastroOrdem(Screen):
    def on_pre_enter(self):
        arq_slvTurma = openpyxl.Workbook()
        arq_slvTurma = openpyxl.load_workbook('cadastro_turmas.xlsx')
        celulas = arq_slvTurma.active
        
        turmas_cadastradas = [celula.value for celula in celulas['A'][1:]]
        num_turmas = len(turmas_cadastradas)
        self.ids.spTurmaEscolhida.values = turmas_cadastradas
        self.ids.spSegunda.values = [str(i) for i in range(1, num_turmas + 1)]
        self.ids.spTerca.values = [str(i) for i in range(1, num_turmas + 1)]
        self.ids.spQuarta.values = [str(i) for i in range(1, num_turmas + 1)]
        self.ids.spQuinta.values = [str(i) for i in range(1, num_turmas + 1)]
        self.ids.spSexta.values = [str(i) for i in range(1, num_turmas + 1)]

    def salvar_Ordem(self, turma, segunda, terca, quarta, quinta, sexta):
        try:
            arq_slvDataOrdem = openpyxl.load_workbook('cadastro_ordemTurmas.xlsx')
        except FileNotFoundError:
            arq_slvDataOrdem = openpyxl.Workbook()
            celula = arq_slvDataOrdem.active
            celula.append(['Turma', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta'])
        
        celula = arq_slvDataOrdem.active
        celula.append([turma, segunda, terca, quarta, quinta, sexta])
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