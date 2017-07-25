VERSION 5.00
Begin VB.Form FRMlogaruser 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sistema de Controle de Lan House"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   6060
      Left            =   0
      Picture         =   "FRMlogaruser.frx":0000
      ScaleHeight     =   6000
      ScaleWidth      =   7500
      TabIndex        =   0
      Top             =   0
      Width           =   7560
      Begin VB.TextBox TXTsenha 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   4785
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2670
         Width           =   2325
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   390
         Left            =   5595
         TabIndex        =   4
         Top             =   3165
         Width           =   1140
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   390
         Left            =   3990
         TabIndex        =   3
         Top             =   3165
         Width           =   1140
      End
      Begin VB.TextBox TXTuser 
         Height          =   345
         Left            =   4785
         TabIndex        =   1
         Top             =   2280
         Width           =   2325
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   7320
         X2              =   2760
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   2760
         X2              =   7320
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   7320
         X2              =   7320
         Y1              =   1800
         Y2              =   3960
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   1
         Left            =   3960
         TabIndex        =   6
         Top             =   2760
         Width           =   705
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         Caption         =   "Nome de usuário:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   3000
         TabIndex        =   5
         Top             =   2280
         Width           =   1725
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   2760
         X2              =   2760
         Y1              =   1800
         Y2              =   3960
      End
   End
End
Attribute VB_Name = "FRMlogaruser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdOK_Click()
Dim banco As Database
Dim dados As Recordset
Dim abrir As String
Dim erro As String
erro = 0
If Right(App.Path, 1) = "\" Then
abrir = App.Path & "base_de_dados.mdb"
Else
abrir = App.Path & "\base_de_dados.mdb"
End If
Set banco = OpenDatabase(abrir)
Set dados = banco.OpenRecordset("funcionarios", dbOpenTable)
Do While dados.EOF = False
If TXTuser.Text = dados.Fields("usuario") Then
If TXTsenha.Text = dados.Fields("senha") Then
If dados.Fields("tipo_user") = "Administrador" Then
admin = 1
erro = 1
FRMcontrolmachine.Show
Unload Me
Else
erro = 1
FRMcontrolmachine.Show
Unload Me
End If
End If
End If
dados.MoveNext
Loop
If erro = 0 Then
MsgBox "Usuário ou senha inválidos", vbCritical, "Falha ao logar"
End If
End Sub

Private Sub Command1_Click()
FRMnovousuario.Show
End Sub

Private Sub Form_Load()
admin = 0
End Sub

Private Sub Picture1_Click()
TXTuser.SetFocus
End Sub
