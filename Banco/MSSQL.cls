VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MSSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private P_cnConexao As ADODB.Connection
Private P_rsInfoTabela As ADODB.Recordset
Private P_rsColunas As ADODB.Recordset
Private P_rsFiltro As ADODB.Recordset
Private P_TransacaoAberta As Boolean
Private P_RecordsetAberto As Boolean
Private P_Tabela As String
Private P_CapturaAutoNumeracao As Boolean
Private P_AutoNumeracao As Long
Private P_TotalRegistros As Long
Private P_Falha As Boolean
Private P_MensagemErro As String
Private P_InstrucaoSql As String

Public Property Let CapturaAutoNumeracao(ByVal Valor As Boolean)
    P_CapturaAutoNumeracao = Valor
End Property

Public Property Get CapturaAutoNumeracao() As Boolean
    CapturaAutoNumeracao = P_CapturaAutoNumeracao
End Property

Public Property Get AutoNumeracao() As Long
    AutoNumeracao = P_AutoNumeracao
End Property

Public Property Get TotalRegistros() As Long
    TotalRegistros = P_TotalRegistros
End Property

Public Property Get Falha() As Boolean
    Falha = P_Falha
End Property

Public Property Get MensagemErro() As String
    MensagemErro = P_MensagemErro
End Property

Public Property Get InstrucaoSql() As String
    InstrucaoSql = P_InstrucaoSql
End Property

'****************************************************************************************************************************

Private Sub Class_Initialize()

    Call Conectar
    
    Set P_rsColunas = CreateObject("ADODB.Recordset")
    P_rsColunas.CursorLocation = adUseClient
    Set P_rsColunas.ActiveConnection = Nothing
    P_rsColunas.Fields.Append "Coluna", adVarChar, 100
    P_rsColunas.Fields.Append "Valor", adVarChar, 8000
    P_rsColunas.Fields.Append "ValidaValor", adBoolean
    P_rsColunas.Open
    
    Set P_rsFiltro = CreateObject("ADODB.Recordset")
    P_rsFiltro.CursorLocation = adUseClient
    Set P_rsFiltro.ActiveConnection = Nothing
    P_rsFiltro.Fields.Append "Coluna", adVarChar, 100
    P_rsFiltro.Fields.Append "Valor", adVarChar, 8000
    P_rsFiltro.Fields.Append "OperadorComparacao", adVarChar, 10
    P_rsFiltro.Fields.Append "OperadorLogico", adVarChar, 5
    P_rsFiltro.Fields.Append "ValidaValor", adBoolean
    P_rsFiltro.Open
    
    P_TransacaoAberta = False
    P_RecordsetAberto = False
    P_Tabela = ""
    P_CapturaAutoNumeracao = False
    P_AutoNumeracao = -1
    P_TotalRegistros = 0
    P_Falha = False
    P_MensagemErro = ""
    P_InstrucaoSql = ""

End Sub

'****************************************************************************************************************************

Private Function Decriptar(ByVal Valor As String) As String

    Dim Cont As Integer
    
    Decriptar = ""
    
    For Cont = 1 To Len(Valor) Step 3
        Decriptar = Decriptar & Chr(Mid(Valor, Cont, 3))
    Next
    
End Function

'****************************************************************************************************************************

Private Sub Conectar()

    Dim Login As String
    Dim Senha As String
    Dim Banco As String
    Dim Servidor As String
    Dim Arquivo As Integer
    Dim Linha As String
    
    Arquivo = FreeFile
    
    Open App.Path & "\ParametersConnection.ini" For Input As #Arquivo
    
    Line Input #Arquivo, Linha
    Login = Decriptar(Linha)
    
    Line Input #Arquivo, Linha
    Senha = Decriptar(Linha)
    
    Line Input #Arquivo, Linha
    Banco = Decriptar(Linha)
    
    Line Input #Arquivo, Linha
    Servidor = Decriptar(Linha)
    
    Close #Arquivo
    
    Set P_cnConexao = CreateObject("ADODB.Connection")
    P_cnConexao.ConnectionString = "Provider=SQLNCLI;Server=" & Servidor & ";Database=" & Banco & ";Uid=" & Login & ";Pwd=" & Senha & ";"
    P_cnConexao.Open
    
End Sub

'****************************************************************************************************************************

Public Sub Desconectar()

    If P_TransacaoAberta Then
        P_cnConexao.CommitTrans
        P_TransacaoAberta = False
    End If
    
    If P_RecordsetAberto Then
        P_rsInfoTabela.Close
        Set P_rsInfoTabela = Nothing
        P_RecordsetAberto = False
    End If
    
    P_rsColunas.Close
    Set P_rsColunas = Nothing
    
    P_rsFiltro.Close
    Set P_rsFiltro = Nothing
    
    P_cnConexao.Close
    Set P_cnConexao = Nothing
    
End Sub

'****************************************************************************************************************************

Public Sub ConfirmarTransacao()

    If P_TransacaoAberta Then
        P_cnConexao.CommitTrans
        P_TransacaoAberta = False
    End If
    
End Sub

'****************************************************************************************************************************

Public Sub ConfigurarTabela(ByVal Valor As String)

    Dim Sql As String
    
    If Not P_RecordsetAberto Then
        Set P_rsInfoTabela = CreateObject("ADODB.Recordset")
        P_rsInfoTabela.CursorLocation = adUseClient
        P_RecordsetAberto = True
    End If
    
    If UCase(Valor) <> UCase(P_Tabela) Then
    
        If P_rsInfoTabela.State = adStateOpen Then
            P_rsInfoTabela.Close
        End If
        
        Sql = "SELECT c.name AS Coluna, UPPER(TYPE_NAME(c.system_type_id)) AS TipoDado, c.is_nullable AS Nulo, CASE WHEN c.default_object_id <> 0 THEN 1 ELSE 0 END AS ValorPadrao FROM sys.objects AS o INNER JOIN sys.columns AS c ON o.object_id = c.object_id WHERE (SCHEMA_NAME(o.schema_id) + '.' + o.name) = '" & Valor & "'"
        P_rsInfoTabela.Open Sql, P_cnConexao
        Set P_rsInfoTabela.ActiveConnection = Nothing
    End If
    
    P_Tabela = Valor

End Sub

'****************************************************************************************************************************

Public Sub AdicionarColuna(ByVal Coluna As String, _
                           ByVal Valor As Variant, _
                           Optional ByVal ValidaValor As Boolean = True)
                            
    If Not IsNull(Valor) Then
        Valor = Trim(CStr(Valor))
        
    Else
        Valor = ""
        
    End If
    
    P_rsColunas.AddNew
    P_rsColunas("Coluna") = Coluna
    P_rsColunas("Valor") = Valor
    P_rsColunas("ValidaValor") = ValidaValor
    
End Sub

'****************************************************************************************************************************

Public Sub AdicionarFiltro(ByVal Coluna As String, _
                           ByVal Valor As Variant, _
                           Optional ByVal OperadorComparacao As String = "=", _
                           Optional ByVal OperadorLogico As String = "AND", _
                           Optional ByVal ValidaValor As Boolean = True)
                            
    If Not IsNull(Valor) Then
        Valor = Trim(CStr(Valor))
        
    Else
        Valor = ""
        
    End If
                            
    P_rsFiltro.AddNew
    P_rsFiltro("Coluna") = Coluna
    P_rsFiltro("Valor") = Valor
    P_rsFiltro("OperadorComparacao") = UCase(OperadorComparacao)
    P_rsFiltro("OperadorLogico") = UCase(OperadorLogico)
    P_rsFiltro("ValidaValor") = ValidaValor
    
End Sub

'****************************************************************************************************************************

Private Sub ExecutarComando(ByVal TipoOperacao As String)

    Dim rsGeral As ADODB.Recordset
    Dim Sql As String
    Dim Sql1 As String
    Dim Sql2 As String
    Dim ColunaValida As Boolean
    
    If P_Falha Then
        Exit Sub
    End If
    
    If Not P_TransacaoAberta Then
        P_cnConexao.BeginTrans
        P_TransacaoAberta = True
    End If
    
On Error GoTo ERRO

    Sql1 = ""
    Sql2 = ""
    ColunaValida = False
    
    If TipoOperacao = "Adicionar" Or TipoOperacao = "Alterar" Then
    
        If P_rsInfoTabela.RecordCount > 0 Then
        
            If P_rsColunas.RecordCount > 0 Then
                P_rsColunas.MoveFirst
                
                While Not P_rsColunas.EOF
                    P_rsInfoTabela.MoveFirst
                    
                    Do While Not P_rsInfoTabela.EOF
                    
                        If UCase(P_rsColunas("Coluna")) = UCase(P_rsInfoTabela("Coluna")) Then
                            P_rsColunas("Valor") = FormatarValor(P_rsColunas("Valor"), , P_rsColunas("ValidaValor"))
                            
                            If Len(P_rsColunas("Valor")) > 0 Then
                            
                                If TipoOperacao = "Adicionar" Then
                                    Sql1 = Sql1 & P_rsColunas("Coluna") & ", "
                                    Sql2 = Sql2 & P_rsColunas("Valor") & ", "
                                    
                                Else
                                    Sql1 = Sql1 & P_rsColunas("Coluna") & " = " & P_rsColunas("Valor") & ", "
                                    
                                End If
                                
                            Else
                                GoTo ERRO
                                
                            End If
                            
                            ColunaValida = True
                            Exit Do
                        End If
                        
                        P_rsInfoTabela.MoveNext
                    Loop
                    
                    If ColunaValida Then
                        ColunaValida = False
                        
                    Else
                        P_MensagemErro = "Módulo: " & TipoOperacao & Chr(13) & "Tabela: " & P_Tabela & Chr(13) & "Coluna: " & P_rsColunas("Coluna") & Chr(13) & "Erro: Coluna inválida"
                        GoTo ERRO
                        
                    End If
                    
                    P_rsColunas.Delete
                    P_rsColunas.MoveNext
                Wend
                
            Else
                P_MensagemErro = "Módulo: " & TipoOperacao & Chr(13) & "Tabela: " & P_Tabela & Chr(13) & "Erro: Nenhuma coluna selecionada"
                GoTo ERRO
                
            End If
                
        Else
            P_MensagemErro = "Módulo: " & TipoOperacao & Chr(13) & "Tabela: " & P_Tabela & Chr(13) & "Erro: Tabela inválida"
            GoTo ERRO
            
        End If
        
        If TipoOperacao = "Adicionar" Then
            Sql1 = "INSERT INTO " & P_Tabela & " (" & Left(Sql1, (Len(Sql1) - 2)) & ")"
            Sql2 = " VALUES (" & Left(Sql2, (Len(Sql2) - 2)) & ")"
            P_InstrucaoSql = Sql1 & Sql2
            P_cnConexao.Execute P_InstrucaoSql
            
            If P_CapturaAutoNumeracao Then
                Sql = "SELECT CASE WHEN SCOPE_IDENTITY() IS NULL THEN -1 ELSE SCOPE_IDENTITY() END AS AutoNumeracao"
                Set rsGeral = CreateObject("ADODB.Recordset")
                rsGeral.CursorLocation = adUseClient
                rsGeral.Open Sql, P_cnConexao
                P_AutoNumeracao = rsGeral("AutoNumeracao")
                rsGeral.Close
                Set rsGeral = Nothing
                P_CapturaAutoNumeracao = False
            End If
            
        Else
            Sql1 = "UPDATE " & P_Tabela & " SET " & Left(Sql1, (Len(Sql1) - 2))
            Sql2 = GerarFiltro(True)
            
            If Len(Sql2) = 0 Then
                GoTo ERRO
            End If
            
            P_InstrucaoSql = Sql1 & Sql2
            P_cnConexao.Execute P_InstrucaoSql
            
        End If
        
    Else
        Sql1 = "DELETE FROM " & P_Tabela
        Sql2 = GerarFiltro(True)
        
        If Len(Sql2) = 0 Then
            GoTo ERRO
        End If
        
        P_InstrucaoSql = Sql1 & Sql2
        P_cnConexao.Execute P_InstrucaoSql
        
    End If
    
    Exit Sub
    
ERRO:

    P_cnConexao.RollbackTrans
    P_TransacaoAberta = False
    P_Falha = True
    
    If Len(P_MensagemErro) = 0 Then
        P_MensagemErro = "Módulo: " & TipoOperacao & Chr(13) & "Erro: " & Err.Description
    End If
    
End Sub

'****************************************************************************************************************************

Public Function FormatarValor(ByVal Valor As String, _
                              Optional ByVal TipoValor As String = "", _
                              Optional ByVal ValidaValor As Boolean = True) As String
                                
On Error GoTo ERRO

    If Len(TipoValor) = 0 Then
    
        If Len(Valor) = 0 Then
        
            If CBool(P_rsInfoTabela("ValorPadrao")) Then
                Valor = "DEFAULT"
                
            ElseIf CBool(P_rsInfoTabela("Nulo")) Then
                Valor = "NULL"
                
            Else
                P_MensagemErro = "Módulo: FormatarValor" & Chr(13) & "Tabela: " & P_Tabela & Chr(13) & "Coluna: " & P_rsInfoTabela("Coluna") & Chr(13) & "Erro: Valor em branco não permitido"
                GoTo ERRO
                
            End If
            
        Else
        
            If P_rsInfoTabela("TipoDado") = "BINARY" Or _
                P_rsInfoTabela("TipoDado") = "VARBINARY" Or _
                P_rsInfoTabela("TipoDado") = "BIT" Or _
                P_rsInfoTabela("TipoDado") = "TINYINT" Or _
                P_rsInfoTabela("TipoDado") = "SMALLINT" Or _
                P_rsInfoTabela("TipoDado") = "INT" Or _
                P_rsInfoTabela("TipoDado") = "BIGINT" Then
                
                If ValidaValor Then
                
                    If Not IsNumeric(Valor) Then
                        P_MensagemErro = "Módulo: FormatarValor" & Chr(13) & "Tabela: " & P_Tabela & Chr(13) & "Coluna: " & P_rsInfoTabela("Coluna") & Chr(13) & "Erro: Número inválido"
                        GoTo ERRO
                    End If
                    
                End If
                
            ElseIf P_rsInfoTabela("TipoDado") = "NUMERIC" Or _
                    P_rsInfoTabela("TipoDado") = "DECIMAL" Or _
                    P_rsInfoTabela("TipoDado") = "SMALLMONEY" Or _
                    P_rsInfoTabela("TipoDado") = "MONEY" Or _
                    P_rsInfoTabela("TipoDado") = "FLOAT" Or _
                    P_rsInfoTabela("TipoDado") = "REAL" Then
                    
                If ValidaValor Then
                
                    If IsNumeric(Valor) Then
                        Valor = Replace(Valor, ".", "")
                        Valor = Replace(Valor, ",", ".")
                        
                    Else
                        P_MensagemErro = "Módulo: FormatarValor" & Chr(13) & "Tabela: " & P_Tabela & Chr(13) & "Coluna: " & P_rsInfoTabela("Coluna") & Chr(13) & "Erro: Número inválido"
                        GoTo ERRO
                        
                    End If
                    
                End If
                    
            ElseIf P_rsInfoTabela("TipoDado") = "DATETIME" Or _
                    P_rsInfoTabela("TipoDado") = "SMALLDATETIME" Then
                    
                If ValidaValor Then
                
                    If IsDate(Valor) Then
                        Valor = "'" & Valor & "'"
                    
                    Else
                        P_MensagemErro = "Módulo: FormatarValor" & Chr(13) & "Tabela: " & P_Tabela & Chr(13) & "Coluna: " & P_rsInfoTabela("Coluna") & Chr(13) & "Erro: Data inválida"
                        GoTo ERRO
                        
                    End If
                    
                End If
                    
            ElseIf P_rsInfoTabela("TipoDado") = "CHAR" Or _
                    P_rsInfoTabela("TipoDado") = "VARCHAR" Or _
                    P_rsInfoTabela("TipoDado") = "TEXT" Or _
                    P_rsInfoTabela("TipoDado") = "NCHAR" Or _
                    P_rsInfoTabela("TipoDado") = "NVARCHAR" Or _
                    P_rsInfoTabela("TipoDado") = "NTEXT" Or _
                    P_rsInfoTabela("TipoDado") = "XML" Then
                    
                Valor = "'" & Replace(Valor, "'", "''") & "'"
                
            Else
            
                If ValidaValor Then
                    P_MensagemErro = "Módulo: FormatarValor" & Chr(13) & "Tabela: " & P_Tabela & Chr(13) & "Coluna: " & P_rsInfoTabela("Coluna") & Chr(13) & "Erro: Tipo de dados desconhecido"
                    GoTo ERRO
                End If
                
            End If
            
        End If
    
    Else
    
        Valor = Trim(Valor)
    
        Select Case UCase(TipoValor)
        
            Case "N" 'Numérico
            
                If IsNumeric(Valor) Then
                    Valor = Replace(Valor, ".", "")
                    Valor = Replace(Valor, ",", ".")
                    
                Else
                    P_MensagemErro = "Módulo: FormatarValor" & Chr(13) & "Erro: Número inválido"
                    GoTo ERRO
                    
                End If
                
            Case "D" 'Data
            
                If Not IsDate(Valor) Then
                    P_MensagemErro = "Módulo: FormatarValor" & Chr(13) & "Erro: Data inválida"
                    GoTo ERRO
                End If
                
            Case "T" 'Texto
                Valor = Replace(Valor, "'", "''")
                
            Case Else
                P_MensagemErro = "Módulo: FormatarValor" & Chr(13) & "Erro: Parâmetro inválido"
                GoTo ERRO
                    
        End Select
        
    End If
    
    FormatarValor = Valor
    
    Exit Function

ERRO:

    FormatarValor = ""
    
    If Len(P_MensagemErro) = 0 Then
        P_MensagemErro = "Módulo: FormatarValor" & Chr(13) & "Erro: " & Err.Description
    End If

End Function

'****************************************************************************************************************************

Public Function GerarFiltro(Optional ByVal UsoInterno As Boolean = False) As String

    Dim ColunaValida As Boolean
    
On Error GoTo ERRO

    GerarFiltro = ""

    If UsoInterno Then
        ColunaValida = False
        
        If P_rsInfoTabela.RecordCount > 0 Then
        
            If P_rsFiltro.RecordCount > 0 Then
                P_rsFiltro.MoveFirst
                
                While Not P_rsFiltro.EOF
                    P_rsInfoTabela.MoveFirst
                    
                    Do While Not P_rsInfoTabela.EOF
                    
                        If UCase(P_rsFiltro("Coluna")) = UCase(P_rsInfoTabela("Coluna")) Then
                        
                            If P_rsFiltro("OperadorLogico") <> "AND" And _
                                P_rsFiltro("OperadorLogico") <> "OR" Then
                                
                                P_MensagemErro = "Módulo: GerarFiltro" & Chr(13) & "Tabela: " & P_Tabela & Chr(13) & "Coluna: " & P_rsInfoTabela("Coluna") & Chr(13) & "Erro: Operador lógico inválido"
                                GoTo ERRO
                                
                            ElseIf P_rsFiltro("OperadorComparacao") <> "=" And _
                                    P_rsFiltro("OperadorComparacao") <> "LIKE" And _
                                    P_rsFiltro("OperadorComparacao") <> "<>" And _
                                    P_rsFiltro("OperadorComparacao") <> "NOT LIKE" And _
                                    P_rsFiltro("OperadorComparacao") <> ">" And _
                                    P_rsFiltro("OperadorComparacao") <> ">=" And _
                                    P_rsFiltro("OperadorComparacao") <> "<" And _
                                    P_rsFiltro("OperadorComparacao") <> "<=" Then
                                    
                                P_MensagemErro = "Módulo: GerarFiltro" & Chr(13) & "Tabela: " & P_Tabela & Chr(13) & "Coluna: " & P_rsInfoTabela("Coluna") & Chr(13) & "Erro: Operador de comparação inválido"
                                GoTo ERRO
                                    
                            End If
                        
                            If Len(P_rsFiltro("Valor")) > 0 Then
                                P_rsFiltro("Valor") = FormatarValor(P_rsFiltro("Valor"), , P_rsFiltro("ValidaValor"))
                                
                                If Len(P_rsFiltro("Valor")) > 0 Then
                                
                                    If Len(GerarFiltro) > 0 Then
                                        GerarFiltro = GerarFiltro & " " & P_rsFiltro("OperadorLogico") & " "
                                        
                                    Else
                                        GerarFiltro = GerarFiltro & " WHERE "
                                    
                                    End If
                                    
                                    GerarFiltro = GerarFiltro & "(" & P_rsFiltro("Coluna") & " " & P_rsFiltro("OperadorComparacao") & " " & P_rsFiltro("Valor") & ")"
                                
                                Else
                                    GoTo ERRO
                                    
                                End If
                                
                            Else
                                P_MensagemErro = "Módulo: GerarFiltro" & Chr(13) & "Tabela: " & P_Tabela & Chr(13) & "Coluna: " & P_rsInfoTabela("Coluna") & Chr(13) & "Erro: Valor em branco não permitido"
                                GoTo ERRO
                
                            End If
                            
                            ColunaValida = True
                            Exit Do
                        End If
                        
                        P_rsInfoTabela.MoveNext
                    Loop
                    
                    If ColunaValida Then
                        ColunaValida = False
                        
                    Else
                        P_MensagemErro = "Módulo: GerarFiltro" & Chr(13) & "Tabela: " & P_Tabela & Chr(13) & "Coluna: " & P_rsFiltro("Coluna") & Chr(13) & "Erro: Coluna inválida"
                        GoTo ERRO
                        
                    End If
                    
                    P_rsFiltro.Delete
                    P_rsFiltro.MoveNext
                Wend
                
            Else
                P_MensagemErro = "Módulo: GerarFiltro" & Chr(13) & "Tabela: " & P_Tabela & Chr(13) & "Erro: Nenhuma coluna selecionada"
                GoTo ERRO
                
            End If
            
        Else
            P_MensagemErro = "Módulo: GerarFiltro" & Chr(13) & "Tabela: " & P_Tabela & Chr(13) & "Erro: Tabela inválida"
            GoTo ERRO
            
        End If
        
    Else
    
        If P_rsFiltro.RecordCount > 0 Then
            P_rsFiltro.MoveFirst
            
            While Not P_rsFiltro.EOF
            
                If Len(GerarFiltro) > 0 Then
                    GerarFiltro = GerarFiltro & " " & P_rsFiltro("OperadorLogico") & " "
                    
                Else
                    GerarFiltro = GerarFiltro & " WHERE "
                    
                End If
                
                GerarFiltro = GerarFiltro & "(" & P_rsFiltro("Coluna") & " " & P_rsFiltro("OperadorComparacao") & " " & P_rsFiltro("Valor") & ")"
            
                P_rsFiltro.Delete
                P_rsFiltro.MoveNext
            Wend
            
        End If
        
    End If
    
    Exit Function

ERRO:

    GerarFiltro = ""
    
    If Len(P_MensagemErro) = 0 Then
        P_MensagemErro = "Módulo: GerarFiltro" & Chr(13) & "Erro: " & Err.Description
    End If

End Function

'****************************************************************************************************************************

Public Sub Adicionar()

    Call ExecutarComando("Adicionar")

End Sub

'****************************************************************************************************************************

Public Sub Alterar()

    Call ExecutarComando("Alterar")

End Sub

'****************************************************************************************************************************

Public Sub Excluir()

    Call ExecutarComando("Excluir")

End Sub

'****************************************************************************************************************************

Public Function Consultar(ByVal Sql As String, _
                          Optional ByVal TamanhoPagina As Long = 10) As ADODB.Recordset

    If P_Falha Then
        Exit Function
    End If
    
On Error GoTo ERRO

    P_InstrucaoSql = Sql

    Set Consultar = CreateObject("ADODB.Recordset")
    Consultar.CursorLocation = adUseClient
    Consultar.PageSize = TamanhoPagina
    Consultar.Open P_InstrucaoSql, P_cnConexao
    Set Consultar.ActiveConnection = Nothing
    
    Exit Function
    
ERRO:

    If P_TransacaoAberta Then
        P_cnConexao.RollbackTrans
        P_TransacaoAberta = False
    End If
    
    P_Falha = True
    
    If Len(P_MensagemErro) = 0 Then
        P_MensagemErro = "Módulo: Consultar" & Chr(13) & "Erro: " & Err.Description
    End If

End Function

'****************************************************************************************************************************

Public Function Consultar_Novo(ByVal Sql As String) As String()

    Dim rsGeral As ADODB.Recordset
    Dim Matriz() As String
    Dim Cont1 As Long
    Dim Cont2 As Integer
    
    If P_Falha Then
        Exit Function
    End If
    
On Error GoTo ERRO

    P_InstrucaoSql = Sql
    
    Set rsGeral = CreateObject("ADODB.Recordset")
    rsGeral.CursorLocation = adUseClient
    rsGeral.Open P_InstrucaoSql, P_cnConexao
    
    If rsGeral.RecordCount > 0 Then
        ReDim Matriz((rsGeral.RecordCount - 1), (rsGeral.Fields.Count - 1))
        Cont1 = -1
        P_TotalRegistros = rsGeral.RecordCount
        
        While Not rsGeral.EOF
            Cont1 = Cont1 + 1
            
            For Cont2 = 0 To (rsGeral.Fields.Count - 1)
            
                If Not IsNull(rsGeral.Fields.Item(Cont2)) Then
                    Matriz(Cont1, Cont2) = rsGeral.Fields.Item(Cont2)
                    
                Else
                    Matriz(Cont1, Cont2) = ""
                    
                End If
            
            Next
            
            rsGeral.MoveNext
        Wend
        
        Consultar_Novo = Matriz
        Erase Matriz
        
    Else
        P_TotalRegistros = 0
        
    End If
    
    rsGeral.Close
    Set rsGeral = Nothing
    
    Exit Function
    
ERRO:

    If P_TransacaoAberta Then
        P_cnConexao.RollbackTrans
        P_TransacaoAberta = False
    End If
    
    P_Falha = True
    Erase Consultar_Novo
    P_TotalRegistros = 0
    
    If Len(P_MensagemErro) = 0 Then
        P_MensagemErro = "Módulo: Consultar_Novo" & Chr(13) & "Erro: " & Err.Description
    End If

End Function
