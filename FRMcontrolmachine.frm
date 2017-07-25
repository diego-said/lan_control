VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FRMcontrolmachine 
   BackColor       =   &H00808080&
   Caption         =   "Controle de computadores"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   570
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "FRMcontrolmachine.frx":0000
   ScaleHeight     =   6690
   ScaleWidth      =   7920
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7200
      Top             =   120
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000000&
      Caption         =   "CHAT - Administrador"
      ForeColor       =   &H00000000&
      Height          =   4335
      Left            =   3960
      TabIndex        =   6
      Top             =   2280
      Width           =   3855
      Begin VB.ListBox LSTchat 
         Height          =   3180
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3615
      End
      Begin VB.CommandButton CMDlimpar 
         Caption         =   "&Limpar"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   3600
         Width           =   3615
      End
      Begin VB.CommandButton CMDenviar 
         Caption         =   "&Enviar"
         Default         =   -1  'True
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   3960
         Width           =   1215
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Index           =   0
         Left            =   120
         Top             =   3840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.ListBox LSTconect 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5970
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Liberar/Bloquear Computador"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   4320
      TabIndex        =   2
      Top             =   1320
      Width           =   3135
      Begin VB.CommandButton cmdbloq 
         Caption         =   "&Bloquear"
         Height          =   300
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   1000
      End
      Begin VB.CommandButton CMDliberar 
         Caption         =   "&Liberar"
         Height          =   300
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1000
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         X1              =   120
         X2              =   4800
         Y1              =   1800
         Y2              =   1800
      End
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HORA ATUAL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5280
      TabIndex        =   11
      Top             =   600
      Width           =   1365
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5280
      TabIndex        =   10
      Top             =   960
      Width           =   75
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   3840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de Computadores Conectados"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3510
   End
End
Attribute VB_Name = "FRMcontrolmachine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ind As Integer
Private Sub cmd_hora_Click()
For ind = 1 To nroClientes
    If LSTconect.Text = ip(ind) Then
        Winsock1(ind).SendData (hora)
        Winsock1(ind).SendData (min)
    End If
Next ind
End Sub
Private Sub cmdbloq_Click()
For ind = 1 To nroClientes
    If LSTconect.Text = ip(ind) Then
        Winsock1(ind).SendData ("1")
    End If
Next ind
End Sub
Private Sub CMDenviar_Click()
lista
End Sub
Private Sub CMDliberar_Click()

For ind = 1 To nroClientes
    If LSTconect.Text = ip(ind) Then
        Winsock1(ind).SendData ("libera")
    End If
Next ind
End Sub
Private Sub CMDlimpar_Click()
LSTchat.Clear
text1.Text = ""
End Sub

Private Sub Form_Load()
nroClientes = 0
Winsock1(0).LocalPort = 1234
Winsock1(0).Listen
Dim banco As Database
Dim dados As Recordset
Dim abrir As String
Dim ind As Single
Dim i As Single
If Right(App.Path, 1) = "\" Then
abrir = App.Path & "base_de_dados.mdb"
Else
abrir = App.Path & "\base_de_dados.mdb"
End If
End Sub
Private Sub Timer1_Timer()
Label2.Caption = Time
End Sub
Private Sub WinSock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If Index = 0 Then
If nroClientes > 0 Then
For x = 1 To nroClientes
If Winsock1(x).State = sckClosed Then
Winsock1(x).Accept (requestID)
ip(x) = Winsock1(x).RemoteHostIP
Exit Sub
End If
Next
End If
nroClientes = nroClientes + 1
Load Winsock1(nroClientes)
ReDim Preserve ip(nroClientes) As String
Winsock1(nroClientes).Accept (requestID)
ip(nroClientes) = Winsock1(nroClientes).RemoteHostIP
LSTconect.AddItem ip(nroClientes)
End If
End Sub
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim dados As String
Winsock1(Index).GetData dados
LSTchat.AddItem " - " & dados & " nro = " & Index & Chr(10)
End Sub
Public Sub lista()
For ind = 1 To nroClientes
    If LSTconect.Text = ip(ind) Then
        Winsock1(ind).SendData (text1.Text)
    End If
Next ind
End Sub
