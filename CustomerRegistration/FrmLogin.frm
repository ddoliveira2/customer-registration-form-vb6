VERSION 5.00
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   2940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnEntrar 
      Caption         =   "Entrar"
      Height          =   345
      Left            =   255
      TabIndex        =   4
      Top             =   1455
      Width           =   2490
   End
   Begin VB.TextBox txtPassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   270
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "123456"
      Top             =   1110
      Width           =   2475
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      Left            =   270
      TabIndex        =   1
      Text            =   "admin"
      Top             =   390
      Width           =   2475
   End
   Begin VB.Label Label2 
      Caption         =   "Senha:"
      Height          =   210
      Left            =   270
      TabIndex        =   2
      Top             =   810
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Usuário:"
      Height          =   210
      Left            =   270
      TabIndex        =   0
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEntrar_Click()
    If UCase(txtUser.Text) = "ADMIN" And UCase(txtPassword.Text) = "123456" Then
        Dim main As New FrmMain
        main.Show
    Else
        MsgBox "Usuário inválido", vbInformation, "Aviso"
    End If
End Sub
