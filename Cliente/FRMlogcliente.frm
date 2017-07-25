VERSION 5.00
Begin VB.Form FRMlogcliente 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   25897.96
   ScaleMode       =   0  'User
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   0
      Picture         =   "FRMlogcliente.frx":0000
      ScaleHeight     =   2250
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   0
      Width           =   4530
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5760
      Top             =   1080
   End
End
Attribute VB_Name = "FRMlogcliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    DisableCtrlAltDelete (True)
End Sub
Private Sub Timer1_Timer()
If Picture2.Visible = True Then
Picture2.Visible = False
Else
Picture2.Visible = True
End If
End Sub
