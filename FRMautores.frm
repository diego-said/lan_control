VERSION 5.00
Begin VB.Form FRMautores 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Autores"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   120
      Picture         =   "FRMautores.frx":0000
      ScaleHeight     =   3000
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   120
      Width           =   6030
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Filipe Mendonça - filipecm@uol.com.br"
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
      Top             =   4920
      Width           =   3795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Diego Said - dsaid@uol.com.br"
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
      Top             =   4560
      Width           =   3030
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "A equipe MONSTER é formada por dois alunos do Colégio Comercial Álvares Penteado e também autores deste programa."
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
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   3600
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6120
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Monster Programadores"
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
      TabIndex        =   1
      Top             =   3240
      Width           =   2325
   End
End
Attribute VB_Name = "FRMautores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
