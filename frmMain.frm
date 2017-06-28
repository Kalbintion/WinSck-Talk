VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Winsock Talker"
   ClientHeight    =   4380
   ClientLeft      =   1995
   ClientTop       =   3585
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6840
      TabIndex        =   11
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txtMsg 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   4080
      Width           =   6855
   End
   Begin VB.TextBox txtMsges 
      Height          =   2295
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   1800
      Width           =   7695
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txtIP 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Text            =   "0.0.0.0"
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Text            =   "80"
      Top             =   720
      Width           =   2535
   End
   Begin VB.OptionButton optJoin 
      BackColor       =   &H00808080&
      Caption         =   "Join"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.OptionButton optHost 
      BackColor       =   &H00808080&
      Caption         =   "Host"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   0
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock winSck 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
      LocalPort       =   80
   End
   Begin VB.Label lblPort 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Port:"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblIP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "IP:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Name:"
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
    winSck.SendData (txtName.Text & ": " & txtMsg.Text)
    AddMessage (txtName.Text & ": " & txtMsg.Text)
    txtMsg.Text = ""
End Sub

Private Sub cmdStart_Click()
    If cmdStart.Caption = "Disconnect" Or cmdStart.Caption = "Connecting..." Then
        winSck.Close
        optHost.Enabled = True
        optJoin.Enabled = True
        txtIP.Enabled = True
        lblIP.Enabled = True
        txtPort.Enabled = True
        lblPort.Enabled = True
        txtName.Enabled = True
        lblName.Enabled = True
        cmdStart.Caption = "Start"
        
        txtMsg.Enabled = False
        cmdSend.Enabled = False
    ElseIf cmdStart.Caption = "Start" Then
        optHost.Enabled = False
        optJoin.Enabled = False
        txtIP.Enabled = False
        lblIP.Enabled = False
        txtPort.Enabled = False
        lblPort.Enabled = False
        txtName.Enabled = False
        lblName.Enabled = False
        cmdStart.Caption = "Connecting..."
        
        AddMessage ("Connecting...")
        
        winSck.LocalPort = txtPort.Text
        winSck.RemotePort = txtPort.Text
        winSck.RemoteHost = txtIP.Text
        
        If optHost.Value = True Then
            winSck.Listen
        ElseIf optJoin.Value = True Then
            winSck.Connect
        End If
        
        txtMsg.Enabled = True
        cmdSend.Enabled = True
    End If
End Sub

Private Sub optHost_Click()
    lblIP.Enabled = False
    txtIP.Enabled = False
End Sub

Private Sub optJoin_Click()
    lblIP.Enabled = True
    txtIP.Enabled = True
End Sub

Private Sub winSck_Close()
    AddMessage ("Disconnected.")
End Sub

Private Sub winSck_ConnectionRequest(ByVal requestID As Long)
    AddMessage ("Request Code: " & requestID)
    If winSck.State = sckListening Then winSck.Close
    winSck.Accept (requestID)
    cmdStart.Caption = "Disconnect"
    AddMessage ("Connected.")
End Sub

Private Sub winSck_DataArrival(ByVal bytesTotal As Long)
    Dim s As String
    winSck.GetData s, vbString
    AddMessage (s)
End Sub

Private Sub winSck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddMessage ("Error: " & Number & " - " & Description)
End Sub

Private Sub AddMessage(msg As String)
    txtMsges.Text = Now & " " & msg & vbCrLf & txtMsges.Text
End Sub
