VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContextSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public conn As ADODB.Connection
Dim rs As ADODB.Recordset

Public Sub ExecutaComandoSql(sql As String)
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=CustomerDb;User ID=sa;Password=123456;"
    conn.Open
    conn.Execute sql
    conn.close
End Sub


Public Function Consulta(ByVal sql As String) As ADODB.Recordset
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=CustomerDb;User ID=sa;Password=123456;"
    Set rs = New ADODB.Recordset
    conn.Open
    rs.Open sql, conn, adOpenStatic, adLockReadOnly

    Set Consulta = rs

End Function



