VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FRMalteraip 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alteração de IP"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FRMalteraip.frx":0000
   ScaleHeight     =   2190
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CMDfechar 
      Caption         =   "F&echar"
      Height          =   300
      Left            =   3840
      TabIndex        =   3
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CommandButton CMDalterar 
      Caption         =   "A&lterar"
      Default         =   -1  'True
      Height          =   300
      Left            =   4920
      TabIndex        =   2
      Top             =   1800
      Width           =   1000
   End
   Begin MSMask.MaskEdBox MSKipatual 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox MSKipnovo 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   1400
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
End
Attribute VB_Name = "FRMalteraip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDalterar_Click()
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
MSKipatual.Mask = ""
MSKipnovo.Mask = ""
If Trim(MSKipnovo.Text) = "___.___.___.___" Then
MsgBox "O campo ip não foi preenchido!", vbCritical, "Campo em branco"
MSKip.Mask = "###.###.###.###"
Else
Do While dados.EOF = False
If MSKipatual.Text = dados.Fields("ip") Then
dados.Edit
dados.Fields("ip") = MSKipnovo.Text
dados.Update
MsgBox "Ip alterado com sucesso!", vbOKOnly, "Ip alterado!"
MSKipnovo.Mask = "###.###.###.###"
End If
dados.MoveNext
Loop
End If
End Sub

Private Sub CMDfechar_Click()
Unload Me
End Sub

Private Sub Form_Load()
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
Do While dados.EOF = False
If ip = dados.Fields("ip") Then
MSKipatual.Mask = ""
MSKipatual.Text = dados.Fields("ip")
MSKipatual.Mask = "###.###.###.###"
End If
dados.MoveNext
Loop
End Sub
