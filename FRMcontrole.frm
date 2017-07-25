VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FRMcontrole 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Liberar/Bloquear estação"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CMDbloquear 
      Caption         =   "Bl&oquear"
      Default         =   -1  'True
      Height          =   390
      Left            =   8280
      TabIndex        =   12
      Top             =   3360
      Width           =   1140
   End
   Begin VB.CommandButton CMDliberar 
      Caption         =   "Li&berar"
      Height          =   390
      Left            =   6360
      TabIndex        =   11
      Top             =   3360
      Width           =   1140
   End
   Begin VB.TextBox TXTuser 
      Height          =   345
      Left            =   7425
      TabIndex        =   8
      Top             =   1320
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   390
      Left            =   7320
      TabIndex        =   7
      Top             =   2205
      Width           =   1140
   End
   Begin VB.TextBox TXTsenha 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   7425
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1710
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "Liberar Computador"
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
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox CMBtempo 
         Height          =   315
         ItemData        =   "FRMcontrole.frx":0000
         Left            =   2040
         List            =   "FRMcontrole.frx":0028
         TabIndex        =   5
         Text            =   "Selecione"
         Top             =   840
         Width           =   1575
      End
      Begin VB.ListBox LSTcompcad 
         Height          =   2400
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   4815
      End
      Begin MSMask.MaskEdBox MSKip 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
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
      Begin VB.Label LBLtempo 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Período de tempo:"
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
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label LBLip 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Ip da estação:"
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
         TabIndex        =   3
         Top             =   360
         Width           =   1470
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   5400
      X2              =   5400
      Y1              =   840
      Y2              =   3000
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Caption         =   "Novo de usuário:"
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
      Index           =   0
      Left            =   5640
      TabIndex        =   10
      Top             =   1320
      Width           =   1680
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
      Left            =   6600
      TabIndex        =   9
      Top             =   1800
      Width           =   705
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   9960
      X2              =   9960
      Y1              =   840
      Y2              =   3000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   5400
      X2              =   9960
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   9960
      X2              =   5400
      Y1              =   3000
      Y2              =   3000
   End
End
Attribute VB_Name = "FRMcontrole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
