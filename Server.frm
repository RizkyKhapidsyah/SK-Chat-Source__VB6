VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   435
      Left            =   3585
      TabIndex        =   3
      Top             =   3405
      Width           =   1215
   End
   Begin VB.TextBox txtReceiverMsg 
      Height          =   1455
      Left            =   1410
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1710
      Width           =   4095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   435
      Left            =   1920
      TabIndex        =   1
      Top             =   3390
      Width           =   1215
   End
   Begin VB.TextBox txtSenderMsg 
      Height          =   1335
      Left            =   1410
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   150
      Width           =   4095
   End
   Begin MSWinsockLib.Winsock wsckServer 
      Left            =   5730
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Message from the Client:"
      Height          =   675
      Left            =   165
      TabIndex        =   5
      Top             =   1695
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Message to the Client:"
      Height          =   675
      Left            =   195
      TabIndex        =   4
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
Dim msg As String
msg = txtSenderMsg
wsckServer.SendData msg
End Sub
Private Sub Form_Activate()
wsckServer.LocalPort = "1001"
wsckServer.Listen
End Sub
Private Sub Form_Terminate()
wsckServer.Close
End Sub
Private Sub wsckServer_ConnectionRequest(ByVal requestID As Long)
If wsckServer.State <> sckClosed Then wsckServer.Close
wsckServer.Accept requestID
End Sub
Private Sub wsckServer_DataArrival(ByVal bytesTotal As Long)
Dim msg As String
wsckServer.GetData msg
txtReceiverMsg = msg
End Sub
