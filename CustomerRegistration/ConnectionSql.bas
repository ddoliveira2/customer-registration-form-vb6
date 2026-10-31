Attribute VB_Name = "ConnectionSql"
Private conn As ADODB.Connection

Public Sub OpenConnection()

    Set conn = New ADODB.Connection

    conn.ConnectionString = ""

    conn.Open

End Sub

Public Sub CloseConnection()

    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            conn.close
        End If

        Set conn = Nothing
    End If

End Sub

Public Property Get Connection() As ADODB.Connection
    Set Connection = conn
End Property

