Attribute VB_Name = "Módulo1"
Sub Primeiro()
'O comando DIM (Dimension) é utilizado para declarar variável
'A Variável Nome foi tipada como String (texto)

Dim Nome As String

'O comando InputBox, abre uma caixa de entrada de dados
'Assim o usuário digita o nome e aloca na variável nome

Nome = InputBox("Digite o seu nome")

'O comando Range, permite selecionar uma célula na planilha do Excel
'Assim selecionamos a célula A1 e adicionamos o valor que foi digitado na caixa de entrada usando a variável Nome

Range("A1").Value = Nome

End Sub
