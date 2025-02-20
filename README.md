<h1>App MóFome</h1>

<h2>Resumo</h2>
O seguinte trabalho está sendo desenvolvido para a matéria de Projeto Interdisciplinar para Sistemas de Informação I, presente na grade curricular do primeiro período do curso de 
<strong>Sistemas de Informação da Universidade Federal Rural de Pernambuco (UFRPE)</strong> e tem como linguagem utilizada o Python junto ao Kivy, framework utilizado para criação de 
aplicativos. A ideia do App surgiu ao se notar a forma com que as escolas que ofertam merenda controlam o processo das filas de alunos no refeitório, sendo bastante bagunçado e quase 
totalmente manual. Portanto, o aplicativo busca resolver essa desorganização. Já a ideia do nome veio da junção das iniciais do tema, surgindo <strong>MóFome</strong>.

<h2>Tema</h2>
Monitoramento de fila de merenda escolar. 

<h2>Objetivo</h2>
Servir como gerenciador de fluxo de fila para a merenda escolar, organizando e controlando os alunos no refeitório da escola. 

<h2>Funcionalidades</h2>
Cadastro de Monitores, Turmas e Ordem na fila; Registro do Dia; Relatório da Semana.

<strong>Cadastro de Monitores:</strong> Nome do aluno monitor, sendo relevante no registro do dia.<br>
<strong>Cadastro de Turmas:</strong> Nome da turma e quantidade de alunos, crucial tanto na tela de Cadastro de Ordem da Fila, como também no Registro do Dia e no Relatório da Semana.<br>
<strong>Cadastro de Ordem na Fila:</strong> Ordena as turmas por dia da semana, importante no registro do dia por turma.<br>
<strong>Registro do Dia:</strong> No registro do dia há a captura dos seguintes dados:<br>
- Data
- Dia da Semana
- Almoço
- Monitor
- Quantidade de alunos por turma e por sexo biológico <br>
<p></p>
<strong>Relatório da Semana:</strong>
O Relatório da Semana apresenta um gráfico de setores e duas tabelas. O gráfico de setores mostra a média de Meninos e Meninas que almoçaram na semana. Aqui ocorre uma média aritmética, onde soma-se os a quantidade de Meninos e Meninas de todas as salas, por dia, e divide por 5 (já que leva-se em conta os dias de segunda-feira a sexta-feira). Disso, é extraído a porcentagem. O relatório mostra também, em tabela, o dia da semana que menos alunos almoçaram e o dia que mais alunos almoçaram. Ainda, em outra tabela, é mostrado o Ranking das 3 turmas que mais almoçaram na semana, levando em consideração a seguinte fórmula: quantidade que almoçou dividido pela quantidade total de alunos da turma. Nesse caso, o cálculo é realizado com base na quantidade de pessoas da turma, não numa média geral. 
