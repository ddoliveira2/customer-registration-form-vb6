VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmCreateUpdateCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2385
   ClientLeft      =   -75
   ClientTop       =   -75
   ClientWidth     =   6510
   Icon            =   "FrmCreateUpdateCustomer.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox mtbCpf 
      Height          =   255
      Left            =   1995
      TabIndex        =   12
      Top             =   1590
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mtbRg 
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   1605
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   423
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton btnLimpar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   3645
      TabIndex        =   10
      Top             =   1965
      Width           =   1335
   End
   Begin VB.CommandButton btnSalvar 
      Caption         =   "Salvar"
      Height          =   300
      Left            =   2235
      TabIndex        =   9
      Top             =   1935
      Width           =   1335
   End
   Begin VB.CommandButton btnSair 
      Caption         =   "Sair"
      Height          =   300
      Left            =   5055
      TabIndex        =   8
      Top             =   1965
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   105
      TabIndex        =   5
      Top             =   1050
      Width           =   6300
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "FrmCreateUpdateCustomer.frx":000C
      Left            =   5175
      List            =   "FrmCreateUpdateCustomer.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   330
      Width           =   1260
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   300
      Left            =   105
      TabIndex        =   0
      Top             =   345
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "RG:"
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   7
      Top             =   1395
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "CPF:"
      Height          =   255
      Index           =   3
      Left            =   2025
      TabIndex        =   6
      Top             =   1395
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Nome:"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   780
      Width           =   1110
   End
   Begin VB.Label Label2 
      Caption         =   "Status:"
      Height          =   210
      Left            =   5160
      TabIndex        =   2
      Top             =   105
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Código:"
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   1110
   End
End
Attribute VB_Name = "FrmCreateUpdateCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sqlConn As New ContextSql
Dim rs As ADODB.Recordset

Private Sub btnLimpar_Click()

    txtName.Text = ""
    mtbRg.Item.Text = ""
    mtbCpf.Item.Text = ""
    txtName.SetFocus
End Sub

Private Sub btnSair_Click()
    Unload Me
End Sub

Private Sub btnSalvar_Click()

    If txtName.Text = "" Then
        MsgBox "Campo nome não informado", vbInformation, "Aviso"
        Return
    End If
    
    If mtbRg.Text = "" Then
        MsgBox "Campo RG não informado", vbInformation, "Aviso"
        Return
    End If
    
    If mtbCpf.Text = "" Then
        MsgBox "Campo CPF não informado", vbInformation, "Aviso"
        Return
    End If
    
    Dim sql As String
    If txtCode.Text <> "" And IsNumeric(txtCode.Text) Then
        Dim status As String
        
        If cboStatus.ListIndex = 0 Then
            status = "1"
        Else
            status = "0"
        End If
        sql = "Update customer set name = '" & UCase(txtName.Text) & "', rg = '" & mtbRg.Text & "', cpf = '" & mtbCpf.Text & "', status =" & status & " where id = " & txtCode.Text
        sqlConn.ExecutaComandoSql (sql)
        MsgBox "Cliente alterado com sucesso!", vbInformation, "sucesso"
        Unload Me
    Else
        sql = "Insert Into customer (name,rg,cpf,status) values ('" & UCase(txtName.Text) & "','" & mtbRg.Text & "','" & mtbCpf.Text & "',1)"
        sqlConn.ExecutaComandoSql (sql)
        MsgBox "Cliente cadastrado com sucesso!", vbInformation, "sucesso"
        Unload Me
    End If
End Sub


Private Sub Form_Activate()
    If txtCode.Text <> "" And IsNumeric(txtCode.Text) Then
        Set rs = sqlConn.Consulta("Select * from customer where id = " & txtCode.Text)
        If rs.RecordCount > 0 Then
            txtCode.Text = rs.Fields("id").Value
            txtName.Text = rs.Fields("name").Value
            mtbRg.Text = rs.Fields("rg").Value
            mtbCpf.Text = rs.Fields("cpf").Value
            If rs.Fields("status").Value = True Then
                cboStatus.ListIndex = 0
            Else
                cboStatus.ListIndex = 1
            End If
        End If
        sqlConn.conn.close
    End If
End Sub

Private Sub mtbCpf_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    
    If KeyAscii = 32 Then Exit Sub
    
    If KeyAscii = 8 Then Exit Sub
    
    KeyAscii = 0

End Sub

Private Sub mtbRg_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    
    If KeyAscii = 32 Then Exit Sub
    
    If KeyAscii = 8 Then Exit Sub
    
    KeyAscii = 0

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 65 And KeyAscii <= 90) Or _
       (KeyAscii >= 97 And KeyAscii <= 122) Then
        Exit Sub
    End If

    If KeyAscii = 32 Then Exit Sub

    If KeyAscii = 8 Then Exit Sub

    KeyAscii = 0
End Sub
