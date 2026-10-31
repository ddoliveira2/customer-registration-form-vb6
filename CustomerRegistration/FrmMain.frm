VERSION 5.00
Begin VB.MDIForm FrmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Pricipal"
   ClientHeight    =   10470
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17835
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu create_update 
      Caption         =   "Cadastrar / Editar"
      Index           =   1
      Begin VB.Menu customer 
         Caption         =   "Cliente"
         Index           =   1
      End
   End
   Begin VB.Menu close 
      Caption         =   "Sair"
      Index           =   1
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_Click(Index As Integer)
    If MsgBox("Deseja sair?", vbQuestion + vbYesNo, "Sair") = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub customer_Click(Index As Integer)
    Dim listCustomer  As New FrmListCustomer
    listCustomer.Show
End Sub

Private Sub MDIForm_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Width = Screen.Width - 30
    Me.Height = Screen.Height - 675
End Sub
