Attribute VB_Name = "M�dulo1"
Sub Primeiro()
'O comando DIM (Dimension) � utilizado para declarar vari�vel
'A Vari�vel Nome foi tipada como String (texto)

Dim Nome As String

'O comando InputBox, abre uma caixa de entrada de dados
'Assim o usu�rio digita o nome e aloca na vari�vel nome

Nome = InputBox("Digite o seu nome")

'O comando Range, permite selecionar uma c�lula na planilha do Excel
'Assim selecionamos a c�lula A1 e adicionamos o valor que foi digitado na caixa de entrada usando a vari�vel Nome

Range("A1").Value = Nome

End Sub
