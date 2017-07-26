VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim Banco As Object
    Set Banco = CreateObject("Banco.MSSQL")
    
    Banco.CapturaAutoNumeracao = True
    
    Banco.Tabela = "dbo.teste_david"
    Banco.AdicionarColuna "Nome", "Edu'ardo"
    Banco.AdicionarColuna "Idade", "25"
    Banco.AdicionarColuna "Valor", "1.000,55"
    Banco.AdicionarColuna "Data", Now
    Banco.Adicionar

    MsgBox (Banco.InstrucaoSql)
    
    Banco.Tabela = "dbo.teste_david"
    Banco.AdicionarColuna "Nome", "Edu'ardo_alt"
    Banco.AdicionarFiltro "Codigo", Banco.AutoNumeracao
    Banco.Alterar

    MsgBox (Banco.InstrucaoSql)

    Banco.Tabela = "dbo.TabPais"
    Banco.AdicionarColuna "NomePais", "Teste"
    Banco.AdicionarColuna "CodigoIdioma", 1
    Banco.AdicionarColuna "Status", 1
    Banco.Adicionar
    
    MsgBox (Banco.InstrucaoSql)
    
    If Banco.Falha Then
        MsgBox (Banco.MensagemErro)
    End If
    
    Banco.Desconectar
    Set Banco = Nothing
    
End Sub
