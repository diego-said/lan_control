VERSION 5.00
Begin VB.Form FRMentrada 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   360
      Top             =   240
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   0
      Picture         =   "FRMentrada.frx":0000
      ScaleHeight     =   3000
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   6030
   End
End
Attribute VB_Name = "FRMentrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
FRMlogaruser.Show
Unload Me
End Sub
