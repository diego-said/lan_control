VERSION 5.00
Begin VB.Form FRMvisualcomp 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selecione o computador para alteração\exclusão"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FRMvisualcomp.frx":0000
   ScaleHeight     =   4455
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CMDlimpar 
      Caption         =   "&Limpar"
      Height          =   300
      Left            =   3840
      TabIndex        =   5
      Top             =   4080
      Width           =   1000
   End
   Begin VB.CommandButton CMDfechar 
      Caption         =   "&Fechar"
      Height          =   300
      Left            =   2760
      TabIndex        =   4
      Top             =   4080
      Width           =   1000
   End
   Begin VB.CommandButton CMDatualizar 
      Caption         =   "A&tualizar"
      Height          =   300
      Left            =   4920
      TabIndex        =   3
      Top             =   4080
      Width           =   1000
   End
   Begin VB.ListBox LSTlistacomp 
      Appearance      =   0  'Flat
      Height          =   3345
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5775
   End
   Begin VB.CommandButton CMDexcluir 
      Caption         =   "E&xcluir"
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   1000
   End
   Begin VB.CommandButton CMDexibir 
      Caption         =   "&Exibir"
      Default         =   -1  'True
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   4080
      Width           =   1000
   End
End
Attribute VB_Name = "FRMvisualcomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDatualizar_Click()
Dim banco As Database
Dim dados As Recordset
Dim abrir As String
If Right(App.Path, 1) = "\" Then
abrir = App.Path & "base_de_dados.mdb"
Else
abrir = App.Path & "\base_de_dados.mdb"
End If
Set banco = OpenDatabase(abrir)
Set dados = banco.OpenRecordset("computadores", dbOpenTable)
Do While dados.EOF = False
LSTlistacomp.AddItem dados.Fields("ip")
dados.MoveNext
Loop
End Sub

Private Sub CMDexcluir_Click()
Dim banco As Database
Dim dados As Recordset
Dim abrir As String
If Right(App.Path, 1) = "\" Then
abrir = App.Path & "base_de_dados.mdb"
Else
abrir = App.Path & "\base_de_dados.mdb"
End If
Set banco = OpenDatabase(abrir)
Set dados = banco.OpenRecordset("computadores", dbOpenTable)
Do While dados.EOF = False
If LSTlistacomp.Text = dados.Fields("ip") Then
dados.Delete
MsgBox "Exclusão efetuada com sucesso.", vbOKOnly, "Exclusão efetuada"
End If
dados.MoveNext
Loop
End Sub

Private Sub CMDexibir_Click()
Dim banco As Database
Dim dados As Recordset
Dim abrir As String
If Right(App.Path, 1) = "\" Then
abrir = App.Path & "base_de_dados.mdb"
Else
abrir = App.Path & "\base_de_dados.mdb"
End If
Set banco = OpenDatabase(abrir)
Set dados = banco.OpenRecordset("computadores", dbOpenTable)
Do While dados.EOF = False
If LSTlistacomp.Text = dados.Fields("ip") Then
ip = dados.Fields("ip")
End If
dados.MoveNext
Loop
FRMalteracomp.Show
End Sub

Private Sub CMDfechar_Click()
Unload Me
End Sub

Private Sub CMDlimpar_Click()
LSTlistacomp.Clear
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
Set dados = banco.OpenRecordset("computadores", dbOpenTable)
Do While dados.EOF = False
LSTlistacomp.AddItem dados.Fields("ip")
dados.MoveNext
Loop
End Sub

