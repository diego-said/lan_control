VERSION 5.00
Begin VB.Form FRMcontrolcomp 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Controle de Computadores"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CMDalterar 
      Caption         =   "ALTERAÇÃO DE COMPUTADOR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton CMDexcluir 
      Caption         =   "EXCLUSÃO DE COMPUTADOR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton CMDlista 
      Caption         =   "LISTA DE COMPUTADORES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton CMDnovo 
      Caption         =   "NOVO COMPUTADOR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label LBLcontrole 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Controle de Computadores"
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
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   3285
   End
End
Attribute VB_Name = "FRMcontrolcomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDalterar_Click()
FRMvisualcomp.Show
End Sub

Private Sub CMDexcluir_Click()
FRMvisualcomp.Show
End Sub

Private Sub CMDlista_Click()
FRMlistacomp.Show
End Sub

Private Sub CMDnovo_Click()
FRMcomputador.Show
End Sub
