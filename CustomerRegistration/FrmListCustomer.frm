VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmListCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Clientes"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   13110
   Begin MSDataGridLib.DataGrid dgvCustomer 
      Height          =   8070
      Left            =   135
      TabIndex        =   6
      Top             =   1050
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   14235
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pesquisar"
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   12855
      Begin VB.TextBox txtCpf 
         Height          =   285
         Left            =   8085
         TabIndex        =   9
         Top             =   435
         Width           =   3330
      End
      Begin VB.TextBox txtRg 
         Height          =   285
         Left            =   6345
         TabIndex        =   8
         Top             =   420
         Width           =   1665
      End
      Begin VB.CommandButton btnNovo 
         Caption         =   "Novo"
         Height          =   375
         Left            =   135
         TabIndex        =   7
         Top             =   330
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   11520
         TabIndex        =   5
         Top             =   345
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   435
         Width           =   4815
      End
      Begin VB.Label Label1 
         Caption         =   "CPF:"
         Height          =   255
         Index           =   3
         Left            =   8280
         TabIndex        =   3
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "RG:"
         Height          =   255
         Index           =   1
         Left            =   6360
         TabIndex        =   2
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nome:"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   1
         Top             =   195
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmListCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim sqlConn As New ContextSql
Private Sub btnNovo_Click()
    Dim register As New FrmCreateUpdateCustomer
    
    register.Show vbModal
    CarregaGrid
End Sub

Private Sub Command1_Click()
    CarregaGrid
End Sub

Private Sub dgvCustomer_DblClick()
    If dgvCustomer.ApproxCount > 0 Then
        Dim update As New FrmCreateUpdateCustomer
        update.txtCode.Text = dgvCustomer.Columns("id").Value
        update.Show vbModal
        CarregaGrid
    End If
End Sub

Private Sub Form_Load()
    CarregaGrid
End Sub

Private Sub CarregaGrid()

    sql = "Select * from customer where name like '%" & txtName.Text & "%' and rg like '%" & txtRg.Text & "%' and cpf like '%" & txtCpf.Text & "%'"
    Set dgvCustomer.DataSource = sqlConn.Consulta(sql)
    
End Sub


Private Sub txtCpf_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    
    If KeyAscii = 32 Then Exit Sub
    
    If KeyAscii = 8 Then Exit Sub
    
    KeyAscii = 0

End Sub

Private Sub txtRg_KeyPress(KeyAscii As Integer)

    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    
    If KeyAscii = 32 Then Exit Sub
    
    If KeyAscii = 8 Then Exit Sub
    
    KeyAscii = 0
End Sub
