VERSION 5.00
Begin VB.Form FRMalterafunc 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alteração de funcionário"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FRMalterafunc.frx":0000
   ScaleHeight     =   4470
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TXTsenha 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox TXTusuario 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   9
      Top             =   2670
      Width           =   2175
   End
   Begin VB.TextBox TXTtel 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   8
      TabIndex        =   8
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox TXTcod_cidade 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox TXTendereco 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      MaxLength       =   150
      TabIndex        =   6
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox TXTnome 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox TXTcodigo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton CMDfechar 
      Caption         =   "&Fechar"
      Height          =   300
      Left            =   2760
      TabIndex        =   3
      Top             =   4080
      Width           =   1000
   End
   Begin VB.CommandButton CMDlimpar 
      Caption         =   "L&impar"
      Height          =   300
      Left            =   3840
      TabIndex        =   2
      Top             =   4080
      Width           =   1000
   End
   Begin VB.CommandButton CMDatualizar 
      Caption         =   "A&tualizar"
      Height          =   300
      Left            =   4920
      TabIndex        =   1
      Top             =   4080
      Width           =   1000
   End
   Begin VB.ComboBox CMBtipouser 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FRMalterafunc.frx":A4A0
      Left            =   960
      List            =   "FRMalterafunc.frx":A4AA
      TabIndex        =   0
      Text            =   "Selecione"
      Top             =   3550
      Width           =   1800
   End
End
Attribute VB_Name = "FRMalterafunc"
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
Set dados = banco.OpenRecordset("funcionarios", dbOpenTable)
If Trim(TXTcodigo.Text) = "" Then
MsgBox "O campo código não foi preenchido!", vbCritical, "Campo em branco"
ElseIf Trim(TXTnome.Text) = "" Then
MsgBox "O campo nome não foi preenchido!", vbCritical, "Campo em branco"
ElseIf Trim(TXTendereco.Text) = "" Then
MsgBox "O campo endereço não foi preenchido!", vbCritical, "Campo em branco"
ElseIf Trim(TXTtel.Text) = "" Then
MsgBox "O campo telefone não foi preenchido!", vbCritical, "Campo em branco"
ElseIf IsNumeric(TXTtel.Text) = False Then
MsgBox "O campo telefone só pode conter números", vbCritical, "Idade inválida"
ElseIf Trim(TXTusuario.Text) = "" Then
MsgBox "O campo apelido não foi preenchido!", vbCritical, "Campo em branco"
ElseIf CMBtipouser.Text = "Selecione" Then
MsgBox "O campo tipo de usuário não foi selecionado, por favor selecione!", vbCritical, "Campo não selecionado"
ElseIf Trim(TXTsenha.Text) = "" Then
MsgBox "O campo senha não foi preenchido!", vbCritical, "Campo em branco"
Else
Do While dados.EOF = False
If TXTcodigo.Text = dados.Fields("codigo") Then
dados.Edit
dados.Fields("codigo") = TXTcodigo.Text
dados.Fields("nome") = TXTnome.Text
dados.Fields("endereco") = TXTendereco.Text
dados.Fields("cod_cidade") = TXTcod_cidade.Text
dados.Fields("telefone") = TXTtel.Text
dados.Fields("usuario") = TXTusuario.Text
dados.Fields("senha") = TXTsenha.Text
dados.Fields("tipo_user") = CMBtipouser.Text
dados.Update
MsgBox "Cadastro alterado com sucesso!", vbOKOnly, "Funcionário alterado!"
End If
dados.MoveNext

Loop
End If
End Sub

Private Sub CMDfechar_Click()
Unload Me
End Sub

Private Sub CMDlimpar_Click()
TXTcodigo.Text = ""
TXTnome.Text = ""
TXTendereco.Text = ""
TXTcod_cidade.Text = ""
TXTtel.Text = ""
TXTusuario.Text = ""
TXTsenha.Text = ""
CMBtipouser.Text = "Selecione"
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
If codfuncionario = dados.Fields("codigo") Then
TXTcodigo.Text = dados.Fields("codigo")
TXTnome.Text = dados.Fields("nome")
TXTendereco.Text = dados.Fields("endereco")
TXTcod_cidade.Text = dados.Fields("cod_cidade")
TXTtel.Text = dados.Fields("telefone")
TXTusuario.Text = dados.Fields("usuario")
TXTsenha.Text = dados.Fields("senha")
CMBtipouser.Text = dados.Fields("tipo_user")
End If
dados.MoveNext
Loop
End Sub

