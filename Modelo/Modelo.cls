VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Modelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public MensagemErro As String

'****************************************************************************************************************************

Public Function Adicionar(ByVal Formulario As Object) As Boolean

    Adicionar = False
    
On Error GoTo ERRO

    Dim Banco As Object
    Set Banco = CreateObject("SysProd.Banco")
    
    Banco.ConfigurarTabela "?"
    Banco.AdicionarColuna "?", Formulario("?")
    Banco.Adicionar
    
    If Banco.Falha = True Then
        MensagemErro = Banco.MensagemErro
        
    Else
        Adicionar = True
    
    End If
    
    Banco.Desconectar
    Set Banco = Nothing
    
    Exit Function
    
ERRO:

    Adicionar = False
    
    If Len(MensagemErro) = 0 Then
        MensagemErro = "Módulo: Adicionar" & Chr(13) & "Erro: " & Err.Description
    End If

End Function

'****************************************************************************************************************************

Public Function Alterar(ByVal ChavePrimaria As Long, ByVal Formulario As Object) As Boolean

    Alterar = False
    
On Error GoTo ERRO

    Dim Banco As Object
    Set Banco = CreateObject("SysProd.Banco")
    
    Banco.ConfigurarTabela "?"
    Banco.AdicionarColuna "?", Formulario("?")
    Banco.AdicionarFiltro "?", ChavePrimaria
    Banco.Alterar
    
    If Banco.Falha = True Then
        MensagemErro = Banco.MensagemErro
        
    Else
        Alterar = True
    
    End If
    
    Banco.Desconectar
    Set Banco = Nothing
    
    Exit Function
    
ERRO:

    Alterar = False
    
    If Len(MensagemErro) = 0 Then
        MensagemErro = "Módulo: Alterar" & Chr(13) & "Erro: " & Err.Description
    End If

End Function

'****************************************************************************************************************************

Public Function Excluir(ByVal ChavePrimaria As Long) As Boolean

    Excluir = False
    
On Error GoTo ERRO

    Dim Banco As Object
    Set Banco = CreateObject("SysProd.Banco")
    
    Banco.ConfigurarTabela "?"
    Banco.AdicionarColuna "Status", "0"
    Banco.AdicionarFiltro "?", ChavePrimaria
    Banco.Alterar
    
    If Banco.Falha = True Then
        MensagemErro = Banco.MensagemErro
        
    Else
        Excluir = True
    
    End If
    
    Banco.Desconectar
    Set Banco = Nothing
    
    Exit Function
    
ERRO:

    Excluir = False
    
    If Len(MensagemErro) = 0 Then
        MensagemErro = "Módulo: Excluir" & Chr(13) & "Erro: " & Err.Description
    End If

End Function

'****************************************************************************************************************************

Public Function VerificarDuplicidade(ByVal Texto As String, Optional ByVal ChavePrimaria As Long = -1) As Boolean

    VerificarDuplicidade = False
    
On Error GoTo ERRO

    Dim Banco As Object
    Set Banco = CreateObject("SysProd.Banco")
    
    Dim rsGeral As ADODB.Recordset
    Dim Sql As String
    
    Sql = "SELECT 1"
    Sql = Sql & " FROM ?"
    Sql = Sql & " WHERE ? = '" & Banco.FormatarValor(Texto, "T") & "'"
    
    If ChavePrimaria <> -1 Then
        Sql = Sql & " AND ? <> " & Banco.FormatarValor(ChavePrimaria, "N")
    End If
    
    Set rsGeral = Banco.Consultar(Sql)
    
    If Banco.Falha = True Then
        MensagemErro = Banco.MensagemErro
        
    Else
    
        If rsGeral.RecordCount = 0 Then
            VerificarDuplicidade = True
        End If
        
        rsGeral.Close
        
    End If
    
    Set rsGeral = Nothing
    
    Banco.Desconectar
    Set Banco = Nothing
    
    Exit Function
    
ERRO:

    VerificarDuplicidade = False
    
    If Len(MensagemErro) = 0 Then
        MensagemErro = "Módulo: VerificarDuplicidade" & Chr(13) & "Erro: " & Err.Description
    End If

End Function

'****************************************************************************************************************************

Public Function ConsultarResumo(ByVal TipoBusca As Long, _
                                ByVal ParametroBusca As String, _
                                ByVal TipoOrdenacao As String, _
                                ByVal TamanhoPagina As Long) As ADODB.Recordset

On Error GoTo ERRO

    Dim Banco As Object
    Set Banco = CreateObject("SysProd.Banco")
    
    Dim Sql As String
    
    Sql = ""
    Sql = Sql & " "
    Sql = Sql & " WHERE 1 = 1"
    
    ParametroBusca = Banco.FormatarValor(ParametroBusca, "T")
    
    If Len(ParametroBusca) > 0 Then
    
        Select Case TipoBusca
            Case 0
                Sql = Sql & " AND ? LIKE '%" & ParametroBusca & "%'"
            Case Else
                Sql = Sql & " AND ? LIKE '%" & ParametroBusca & "%'"
        End Select
        
    End If
    
    Select Case TipoOrdenacao
        Case "?"
            Sql = Sql & " ORDER BY ? ASC"
        Case Else
            Sql = Sql & " ORDER BY ? DESC"
    End Select
    
    Set ConsultarResumo = Banco.Consultar(Sql, TamanhoPagina)
    
    If Banco.Falha = True Then
        MensagemErro = Banco.MensagemErro
    End If
    
    Banco.Desconectar
    Set Banco = Nothing
    
    Exit Function
    
ERRO:

    If Len(MensagemErro) = 0 Then
        MensagemErro = "Módulo: ConsultarResumo" & Chr(13) & "Erro: " & Err.Description
    End If

End Function

'****************************************************************************************************************************

Public Function ConsultarDetalhes(Optional ByVal ChavePrimaria As Long = -1, Optional ByVal Status As Integer = -1) As ADODB.Recordset

On Error GoTo ERRO

    Dim Banco As Object
    Set Banco = CreateObject("SysProd.Banco")
    
    Dim Sql As String
    
    Sql = ""
    Sql = Sql & " "
    Sql = Sql & " WHERE 1 = 1"
    
    If ChavePrimaria <> -1 Then
        Sql = Sql & " AND ? = " & Banco.FormatarValor(ChavePrimaria, "N")
    End If
    
    If Status <> -1 Then
        Sql = Sql & " AND ? = " & Banco.FormatarValor(Status, "N")
    End If
    
    Sql = Sql & " ORDER BY ? ASC"
    
    Set ConsultarDetalhes = Banco.Consultar(Sql)
    
    If Banco.Falha = True Then
        MensagemErro = Banco.MensagemErro
    End If
    
    Banco.Desconectar
    Set Banco = Nothing
    
    Exit Function
    
ERRO:

    If Len(MensagemErro) = 0 Then
        MensagemErro = "Módulo: ConsultarDetalhes" & Chr(13) & "Erro: " & Err.Description
    End If
    
End Function
