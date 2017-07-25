VERSION 5.00
Begin VB.Form FRMlistafunc 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Funcionários Cadastrados"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FRMlistafunc.frx":0000
   ScaleHeight     =   4470
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CMDlimpar 
      Caption         =   "&Limpar"
      Height          =   300
      Left            =   3600
      TabIndex        =   3
      Top             =   4080
      Width           =   1000
   End
   Begin VB.CommandButton CMDfechar 
      Caption         =   "&Fechar"
      Height          =   300
      Left            =   2520
      TabIndex        =   2
      Top             =   4080
      Width           =   1000
   End
   Begin VB.CommandButton CMDatualizar 
      Caption         =   "A&tualizar"
      Default         =   -1  'True
      Height          =   300
      Left            =   4680
      TabIndex        =   1
      Top             =   4080
      Width           =   1000
   End
   Begin VB.ListBox LSTlistafunc 
      Appearance      =   0  'Flat
      Height          =   3345
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5655
   End
End
Attribute VB_Name = "FRMlistafunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDatualizar_Click()
LSTlistafunc.Clear
Dim banco As Database
Dim dados As Recordset
Dim abrir As String
If Right(App.Path, 1) = "\" Then
abrir = App.Path & "base_de_dados.mdb"
Else
abrir = App.Path & "\base_de_dados.mdb"
End If
Set banco = OpenDatabase(abrir)
Set dados = banco.OpenRecordset("funcionarios", dbOpenTable)
Do While dados.EOF = False
LSTlistafunc.AddItem dados.Fields("nome")
dados.MoveNext
Loop
End Sub

Private Sub CMDfechar_Click()
Unload Me
End Sub

Private Sub CMDlimpar_Click()
LSTlistafunc.Clear
End Sub

Private Sub Form_Load()
Dim banco As Database
Dim dados As Recordset
Dim abrir As String
If Right(App.Path, 1) = "\" Then
abrir = App.Path & "base_de_dados.mdb"
Else
abrir = App.Path & "\base_de_dados.mdb"
End If
Set banco = OpenDatabase(abrir)
Set dados = banco.OpenRecordset("funcionarios", dbOpenTable)
Do While dados.EOF = False
LSTlistafunc.AddItem dados.Fields("nome")
dados.MoveNext
Loop
End Sub
