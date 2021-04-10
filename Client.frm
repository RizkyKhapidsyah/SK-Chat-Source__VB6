VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   4110
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3780
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtReceiverMsg 
      Height          =   1335
      Left            =   1725
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1935
      Width           =   3975
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   2250
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtSenderMsg 
      Height          =   1455
      Left            =   1725
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   255
      Width           =   3975
   End
   Begin MSWinsockLib.Winsock wsckClient 
      Left            =   5805
      Top             =   255
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Message from the Server:"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   1935
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Message to the Server:"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   270
      Width           =   1080
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub optConnect_Click()
End Sub
Private Sub optDisConnect_Click()
wsckClient.Close
optConnect.Enabled = True
optDisconnect.Enabled = False
txtSenderMsg.Enabled = False: txtReceiverMsg.Enabled = False
End Sub
Private Sub Form_Activate()
wsckClient.RemoteHost = "192.168.1.138"    'Edit the Server Machine IP Address as you want
wsckClient.RemotePort = "1001"           'Port Number should be similar as defined in the Server Program
wsckClient.Connect
End Sub
Private Sub wsckClient_DataArrival(ByVal bytesTotal As Long)
Dim msg As String
wsckClient.GetData msg
txtReceiverMsg = msg
End Sub
Private Sub cmdSend_Click()
On Error GoTo last
wsckClient.SendData txtSenderMsg
Exit Sub
last:
If Err.Number = 40006 Then
   MsgBox "Not connected, Try again"
End If
End Sub
Private Sub Form_Terminate()
wsckClient.Close
End Sub

