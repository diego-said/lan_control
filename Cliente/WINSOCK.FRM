VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form winsock 
   BorderStyle     =   0  'None
   Caption         =   "winsock"
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   435
   ScaleWidth      =   465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "winsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FRMlogcliente.Show
winsock.Hide
Winsock1.RemoteHost = "127.0.0.1"
Winsock1.RemotePort = 1234
Winsock1.Connect
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim mensagem As String
Winsock1.GetData mensagem
If mensagem = "1" Then
Unload FRMlogado
FRMlogcliente.Show
ElseIf mensagem = "libera" Then
Unload FRMlogcliente
FRMlogado.Show
Else
FRMlogado.LSTchat.AddItem "Servidor:" & " " & mensagem
End If
End Sub

