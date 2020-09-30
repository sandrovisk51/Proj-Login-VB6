Attribute VB_Name = "modConexao"
Option Explicit
Global bd As ADODB.Connection
Global Comd As ADODB.Command
Global rs As ADODB.Recordset

Public Function Conecta() As Boolean
Set bd = New ADODB.Connection
Set Comd = New ADODB.Command

On Error GoTo base

bd.ConnectionTimeout = 60
bd.CommandTimeout = 100
bd.CursorLocation = 1

bd.Open "Driver=MySQL ODBC 5.1 Driver; Server=localhost; Database=dblogin; uid=root; pwd=eqroot; Port = 3306;"
Set Comd.ActiveConnection = bd
     
Exit Function
base:
    MsgBox "Erro número : " & (Err.Number) & "  --> Verificar a Conexão com a base de dados..."
    Err.Clear
    Exit Function
End Function

Public Function RecosetMysql(query As String)
On Error GoTo recosetErro

Conecta
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open query, bd, adOpenDynamic, adLockOptimistic, adCmdText

Exit Function
recosetErro:
 
    MsgBox "Erro número : " & (Err.Number) & "  --> Erro no processo da transação dos dados."
    Err.Clear
    Exit Function
End Function
