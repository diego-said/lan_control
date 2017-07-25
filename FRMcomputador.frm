VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FRMcomputador 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Computadores"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FRMcomputador.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CMDlimpar 
      Caption         =   "&Limpar"
      Height          =   300
      Left            =   3720
      TabIndex        =   6
      Top             =   4080
      Width           =   1000
   End
   Begin VB.CommandButton CMDfechar 
      Caption         =   "F&echar"
      Height          =   300
      Left            =   2640
      TabIndex        =   5
      Top             =   4080
      Width           =   1000
   End
   Begin VB.CommandButton CMDcadastrar 
      Caption         =   "&Cadastrar"
      Default         =   -1  'True
      Height          =   300
      Left            =   4800
      TabIndex        =   4
      Top             =   4080
      Width           =   1000
   End
   Begin VB.ComboBox CMBmanutencao 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FRMcomputador.frx":9770
      Left            =   1800
      List            =   "FRMcomputador.frx":977A
      TabIndex        =   3
      Text            =   "Selecione"
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox TXTgrupo 
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
      Height          =   300
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   2
      Top             =   2595
      Width           =   2415
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
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin MSMask.MaskEdBox MSKip 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   1280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "###.###.###.###"
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "FRMcomputador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDcadastrar_Click()
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
MSKip.Mask = ""
If Trim(MSKip.Text) = "___.___.___.___" Then
MsgBox "O campo ip não foi preenchido!", vbCritical, "Campo em branco"
MSKip.Mask = "###.###.###.###"
ElseIf Trim(TXTnome.Text) = "" Then
MsgBox "O campo nome não foi preenchido!", vbCritical, "Campo em branco"
ElseIf Trim(TXTgrupo.Text) = "" Then
MsgBox "O campo grupo não foi preenchido!", vbCritical, "Campo em branco"
ElseIf CMBmanutencao.Text = "Selecione" Then
MsgBox "O campo manutenção não foi selecionado, por favor selecione!", vbCritical, "Campo não selecionado"
Else
dados.AddNew
dados.Fields("ip") = MSKip.Text
dados.Fields("nome") = TXTnome.Text
dados.Fields("grupo") = TXTgrupo.Text
dados.Fields("manutencao") = CMBmanutencao.Text
dados.Update
MSKip.Text = ""
MSKip.Mask = "###.###.###.###"
TXTnome.Text = ""
TXTgrupo.Text = ""
CMBmanutencao.Text = "Selecione"
MsgBox "Cadastro efetuado com sucesso!", vbOKOnly, "Computador cadastrado!"
End If
End Sub

Private Sub CMDfechar_Click()
Unload Me
End Sub

Private Sub CMDlimpar_Click()
MSKip.Mask = ""
MSKip.Text = ""
MSKip.Mask = "###.###.###.###"
TXTnome.Text = ""
TXTgrupo.Text = ""
CMBmanutencao.Text = "Selecione"
End Sub
