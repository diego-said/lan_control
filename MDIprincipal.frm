VERSION 5.00
Begin VB.MDIForm MDIprincipal 
   BackColor       =   &H00808080&
   Caption         =   "Controle de LAN House - versão 1.0"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu MNUclientes 
      Caption         =   "&Clientes"
      Begin VB.Menu MNUnovocliente 
         Caption         =   "&Novo"
         Shortcut        =   ^B
      End
      Begin VB.Menu MNUlistacliente 
         Caption         =   "&Lista"
         Shortcut        =   ^C
      End
      Begin VB.Menu MNUaltexccliente 
         Caption         =   "&Alterar/Excluir"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu MNUfuncionarios 
      Caption         =   "&Funcionários"
      Begin VB.Menu MNUnovofuncionario 
         Caption         =   "&Novo"
         Shortcut        =   ^E
      End
      Begin VB.Menu MNUlistafuncionario 
         Caption         =   "&Lista"
         Shortcut        =   ^F
      End
      Begin VB.Menu MNUaltexcfuncionario 
         Caption         =   "&Alterar/Excluir"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu MNUcomputadores 
      Caption         =   "C&omputadores"
      Begin VB.Menu MNUnovo 
         Caption         =   "&Novo"
         Shortcut        =   ^H
      End
      Begin VB.Menu MNUlista 
         Caption         =   "&Lista"
         Shortcut        =   ^I
      End
      Begin VB.Menu MNUaltexccomputador 
         Caption         =   "&Alterar/Excluir"
         Shortcut        =   ^J
      End
   End
   Begin VB.Menu MNUajuda 
      Caption         =   "&Ajuda"
      Begin VB.Menu MNUguia 
         Caption         =   "&Guia Completo"
         Shortcut        =   ^A
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu MNUautores 
         Caption         =   "Au&tores"
      End
      Begin VB.Menu MNUsobre 
         Caption         =   "&Sobre o controle de LAN house"
      End
   End
End
Attribute VB_Name = "MDIprincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
'If admin = 0 Then
'FRMfuncionarios.Enabled = False
'FRMcomputador.Enabled = False
'End If
End Sub

Private Sub MNUaltexccliente_Click()
FRMvisualjoga.Show
End Sub

Private Sub MNUaltexccomputador_Click()
FRMvisualcomp.Show
End Sub

Private Sub MNUaltexcfuncionario_Click()
FRMvisualfunc.Show
End Sub

Private Sub MNUautores_Click()
FRMautores.Show
End Sub

Private Sub MNUlista_Click()
FRMlistacomp.Show
End Sub

Private Sub MNUlistacliente_Click()
FRMlistajoga.Show
End Sub

Private Sub MNUlistafuncionario_Click()
FRMlistafunc.Show
End Sub

Private Sub MNUnovo_Click()
FRMcomputador.Show
End Sub

Private Sub MNUnovocliente_Click()
FRMjogadores.Show
End Sub

Private Sub MNUnovofuncionario_Click()
FRMfuncionarios.Show
End Sub

Private Sub MNUsobre_Click()
MsgBox "Este programa foi desenvolvido no ano de 2003, pelos alunos do Colégio Comercial Álvares Penteado, como parte do trabalho de conclusão de curso. Sendo proibida sua reprodução sem autorização prévia.", vbOKOnly, "Sobre o Controle de Lan House"
End Sub
