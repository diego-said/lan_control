VERSION 5.00
Begin VB.Form FRMlogado 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7245
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FRMlogado.frx":0000
   ScaleHeight     =   4845
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   4320
   End
   Begin VB.CommandButton CMDlimpar 
      Caption         =   "&Limpar"
      Height          =   300
      Left            =   4320
      TabIndex        =   3
      Top             =   4440
      Width           =   1000
   End
   Begin VB.TextBox TXTmensagem 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      Top             =   4080
      Width           =   3135
   End
   Begin VB.CommandButton CMDenviar 
      Caption         =   "&Enviar"
      Default         =   -1  'True
      Height          =   300
      Left            =   5880
      TabIndex        =   2
      Top             =   4440
      Width           =   1000
   End
   Begin VB.ListBox LSTchat 
      Appearance      =   0  'Flat
      Height          =   3930
      Left            =   3960
      TabIndex        =   12
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label LBLtotal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "R$ 0,00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2520
      TabIndex        =   6
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label LBLtotpag 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total a pagar:"
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
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label LBL1hora 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1 hora .......................... R$ 3,00"
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
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   3480
   End
   Begin VB.Label LBL45min 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "45 min .......................... R$ 2,25"
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
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label LBL30min 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "30 min .......................... R$ 1,50"
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
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label LBL15min 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "15 min .......................... R$ 0,75"
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
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label LBLprecos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tabela de preços"
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
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   1710
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   3600
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line2 
      X1              =   3840
      X2              =   3840
      Y1              =   120
      Y2              =   4680
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3600
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo conectado:"
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
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   1875
   End
   Begin VB.Label LBLhorac 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2280
      TabIndex        =   10
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label LBLminc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2760
      TabIndex        =   9
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label LBLsegc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3240
      TabIndex        =   5
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label dp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   2640
      TabIndex        =   8
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label dp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   3120
      TabIndex        =   7
      Top             =   1080
      Width           =   105
   End
End
Attribute VB_Name = "FRMlogado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDenviar_Click()
winsock.Winsock1.SendData (TXTmensagem.Text)
LSTchat.AddItem "Você:" & " " & TXTmensagem.Text
End Sub

Private Sub CMDlimpar_Click()
TXTmensagem.Text = ""
End Sub

Private Sub Form_Load()
DisableCtrlAltDelete (True)
End Sub

Private Sub Timer1_Timer()
Dim vmin As Single
Dim vhora As Single
Dim soma As Single
Dim titulo As String
Do While LBLsegc.Caption = 59
LBLsegc.Caption = "00"
If LBLminc.Caption < 59 Then
LBLminc.Caption = LBLminc.Caption + 1
If LBLminc.Caption < 10 Then
LBLminc.Caption = "0" + LBLminc.Caption
End If
Else
LBLminc.Caption = "00"
LBLhorac.Caption = LBLhorac.Caption + 1
If LBLhorac.Caption < 10 Then
LBLhorac.Caption = "0" + LBLhorac.Caption
End If
End If
Loop
If LBLsegc.Caption < 59 Then
LBLsegc.Caption = LBLsegc.Caption + 1
If LBLsegc.Caption < 10 Then
LBLsegc.Caption = "0" + LBLsegc.Caption
End If
End If
titulo = LBLhorac.Caption & ":" & LBLminc.Caption & ":" & LBLsegc.Caption
FRMlogado.Caption = titulo
vmin = (Val(LBLminc.Caption) * 3) / 60
vhora = Val(LBLhorac.Caption) * 3
soma = vhora + vmin
LBLtotal.Caption = "R$ " & soma
End Sub


