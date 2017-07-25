VERSION 5.00
Begin VB.Form FRMjogadores 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Jogadores"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FRMjogadores.frx":0000
   ScaleHeight     =   4440
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CMDlimpar 
      Appearance      =   0  'Flat
      Caption         =   "&Limpar"
      Height          =   300
      Left            =   3840
      TabIndex        =   10
      Top             =   4080
      Width           =   1000
   End
   Begin VB.CommandButton CMDfechar 
      Caption         =   "&Fechar"
      Height          =   300
      Left            =   2760
      Picture         =   "FRMjogadores.frx":ADAE
      TabIndex        =   9
      Top             =   4080
      Width           =   1000
   End
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
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox TXTapelido 
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
      TabIndex        =   7
      Top             =   3120
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
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
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
      TabIndex        =   5
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox TXTidade 
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
      Height          =   270
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   4
      Top             =   2200
      Width           =   375
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
      TabIndex        =   3
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
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1320
      Width           =   4695
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
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton CMDcadastrar 
      Caption         =   "&Salvar"
      Default         =   -1  'True
      Height          =   300
      Left            =   4920
      Picture         =   "FRMjogadores.frx":B41F
      TabIndex        =   0
      ToolTipText     =   "Salvar Cadastro"
      Top             =   4080
      Width           =   1000
   End
End
Attribute VB_Name = "FRMjogadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDlimpar_Click()
TXTnome.Text = ""
TXTendereco.Text = ""
TXTidade.Text = ""
TXTcod_cidade.Text = "011"
TXTtel.Text = ""
TXTapelido.Text = ""
TXTsenha.Text = ""
End Sub
Private Sub CMDfechar_Click()
Unload Me
End Sub

Private Sub CMDcadastrar_Click()
Dim banco As Database
Dim dados As Recordset
Dim abrir As String
Dim codigo As Single
If Right(App.Path, 1) = "\" Then
abrir = App.Path & "base_de_dados.mdb"
Else
abrir = App.Path & "\base_de_dados.mdb"
End If
Set banco = OpenDatabase(abrir)
Set dados = banco.OpenRecordset("jogadores", dbOpenTable)
Do While dados.EOF = False
If TXTcodigo.Text = dados.Fields("codigo") Then
MsgBox "Esse código não pode ser usado, porque já existe.", vbCritical, "Duplicidade de código"
End If
dados.MoveNext
Loop
If Trim(TXTcodigo.Text) = "" Then
MsgBox "O campo código não foi preenchido!", vbCritical, "Campo em branco"
ElseIf Trim(TXTnome.Text) = "" Then
MsgBox "O campo nome não foi preenchido!", vbCritical, "Campo em branco"
ElseIf Trim(TXTendereco.Text) = "" Then
MsgBox "O campo endereço não foi preenchido!", vbCritical, "Campo em branco"
ElseIf Trim(TXTidade.Text) = "" Then
MsgBox "O campo idade não foi preenchido!", vbCritical, "Campo em branco"
ElseIf IsNumeric(TXTidade.Text) = False Then
MsgBox "O campo idade só pode conter números", vbCritical, "Idade inválida"
ElseIf Trim(TXTtel.Text) = "" Then
MsgBox "O campo telefone não foi preenchido!", vbCritical, "Campo em branco"
ElseIf IsNumeric(TXTtel.Text) = False Then
MsgBox "O campo telefone só pode conter números", vbCritical, "Idade inválida"
ElseIf Trim(TXTapelido.Text) = "" Then
MsgBox "O campo apelido não foi preenchido!", vbCritical, "Campo em branco"
ElseIf Trim(TXTsenha.Text) = "" Then
MsgBox "O campo senha não foi preenchido!", vbCritical, "Campo em branco"
Else
dados.AddNew
dados.Fields("codigo") = TXTcodigo.Text
dados.Fields("nome") = TXTnome.Text
dados.Fields("endereco") = TXTendereco.Text
dados.Fields("idade") = TXTidade.Text
dados.Fields("cod_cidade") = TXTcod_cidade.Text
dados.Fields("telefone") = TXTtel.Text
dados.Fields("apelido") = TXTapelido.Text
dados.Fields("senha") = TXTapelido.Text
dados.Update
TXTcodigo.Text = ""
TXTnome.Text = ""
TXTendereco.Text = ""
TXTidade.Text = ""
TXTcod_cidade.Text = ""
TXTtel.Text = ""
TXTapelido.Text = ""
TXTsenha.Text = ""
MsgBox "Cadastro efetuado com sucesso!", vbOKOnly, "Cliente cadastrado!"
Unload Me
FRMjogadores.Show
End If
End Sub

Private Sub Form_Load()
Dim banco As Database
Dim dados As Recordset
Dim abrir As String
Dim codigo As Single
If Right(App.Path, 1) = "\" Then
abrir = App.Path & "base_de_dados.mdb"
Else
abrir = App.Path & "\base_de_dados.mdb"
End If
Set banco = OpenDatabase(abrir)
Set dados = banco.OpenRecordset("jogadores", dbOpenTable)
Do While dados.EOF = False
codigo = codigo + 1
dados.MoveNext
Loop
TXTcodigo.Text = codigo
TXTcod_cidade.Text = "011"
End Sub

