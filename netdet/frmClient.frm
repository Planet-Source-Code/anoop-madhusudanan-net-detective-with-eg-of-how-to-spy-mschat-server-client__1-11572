VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClient 
   Caption         =   "Net Detective"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7860
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fr 
      Height          =   4095
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   7785
      Begin VB.ListBox lstChat 
         Height          =   3375
         Left            =   120
         TabIndex        =   2
         Top             =   660
         Width           =   3705
      End
      Begin VB.ListBox lstServer 
         Height          =   3375
         Left            =   3945
         TabIndex        =   1
         Top             =   675
         Width           =   3705
      End
      Begin MSComctlLib.ImageList imlMain 
         Left            =   2955
         Top             =   30
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":0896
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":0CEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":113E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Data From Client To Server"
         Height          =   360
         Left            =   180
         TabIndex        =   4
         Top             =   315
         Width           =   2415
      End
      Begin VB.Label lblServer 
         Caption         =   "Data From Server To Client"
         Height          =   360
         Left            =   4005
         TabIndex        =   3
         Top             =   300
         Width           =   2415
      End
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   1058
      ButtonWidth     =   2566
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Start Listen"
            Key             =   "Start"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Disconnect"
            Key             =   "Disc"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save Log"
            Key             =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "About"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Frame frSub 
      Height          =   600
      Left            =   15
      TabIndex        =   5
      Top             =   4800
      Width           =   7815
      Begin VB.TextBox txtMessage 
         Height          =   345
         Left            =   930
         TabIndex        =   8
         Top             =   165
         Width           =   4005
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<< To Client"
         Height          =   375
         Left            =   5130
         TabIndex        =   7
         Top             =   135
         Width           =   1125
      End
      Begin VB.CommandButton Command2 
         Caption         =   "To Server >>"
         Height          =   345
         Left            =   6450
         TabIndex        =   6
         Top             =   150
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Message:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   225
         Width           =   690
      End
   End
   Begin MSWinsockLib.Winsock wsListen 
      Left            =   180
      Top             =   645
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsConnect 
      Left            =   4890
      Top             =   660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsAccept 
      Left            =   795
      Top             =   645
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
'INTERNET DETECTIVE - For Spying Data
'================================================
'
'By ANOOP.M, Web Strategist & Developer
'Visit http://profiles.guru.com/anoopm
'
'Send Personal Mail to anoopj13@yahoo.com
'if you need a (quicker) reply
'
'================================================
'Category:
'================================================
'Internet/Sockets
'
'================================================
'Purpose:
'================================================
'For extracting the data transmission between
'two sockets.
'
'================================================
'Help:
'================================================
'Our purpose is to spy the data transmission
'between two sockets. That is, our spy should
'act between the server and client.
'
' Server <-----> Spy <-----> Client
'                 |
'                 |
'                 V
'             Record Data
'
'For example, if you are spying the data trans
'mission between Microsoft Chat Server/Client,
'here is the procedure.
'
'1) After going online, run the App.
'2) If you do not have a permenant IP,
'   get ur current IP from the debug window.
'3) If you are using MS Chat 2.5, goto
'   View->Options->Servers and setup a
'   server with the current IP
'   (For earlier versions of MSChat, I think
'    you can paste it directly in the initial
'    'Chat Connection' box)
'4) Then click 'connect' to connect the Microsoft
'   Chat to your IP(not the IP of MS Chat
'   Server)
'5) Just examine the listbox (better replace with
'   a text box) to see the 'real' data.
'6) Goto your nearby shop and buy a little
'   glucose (for getting energy to understand
'   what you see..
'
'
'   Hope this may help..Download my other apps
'   including a little famous Icon Hunter
'   from the Planet..regards.
'
'
'   Once the client establishes a connection
'   our prog will automatically contact the
'   server..
'
'================================================
'Details:
'================================================
'Well, This is a simple program. If the name
'is doesn't indicate anything, here is a short
'story:
'
'I usually use Microsoft Chat for chatting.
'(Usually my nick name is Nice-Guy, in case
'u need to find me..) Just for understanding
'the communication between the Microsoft chat
'client and the server, I wrote this app..
'
'But you can use it as a router,multi
'chat enabler etc, if
'you need to use it that way..
'================================================

Dim mCurReqId
Dim YesConnect As Boolean

'The server name. In this exmp, it is MS Chat server
Const gServerIp = "mschat1.msn.com"

'The Port of Server..This port for Ms chat server
Const gServerPort = 6667

'The Port of Client..MSChat client requests this port
'So it is the listening port
Const gListenPort = 6667

Private Sub Command1_Click()
'Overrides server to send data to the client
On Error Resume Next
wsAccept.SendData txtMessage.Text + vbCrLf
End Sub

Private Sub Command2_Click()
'Overrides client to send data to the server
On Error Resume Next
wsConnect.SendData txtMessage.Text + vbCrLf
End Sub

Private Sub Form_Load()

'wsListen is the listening socket (for client)
'wsAccept is the socket connected to client
'wsConnect is the scoket connected to server

'Set The variable
YesConnect = False

'For telling you your localip
Debug.Print wsListen.LocalIP
Me.Caption = "Internet Detective - [" & wsListen.LocalIP & "]"

End Sub

Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
'Set enabling/disabling stuff urself..u too need some job..lol

Select Case Button.Key

Case "Start"
    Err.Clear
    On Error Resume Next
    wsListen.LocalPort = gListenPort
    wsListen.Listen

    If Err Then
        MsgBox "Error: " & Err.Description, vbCritical + vbOKOnly
    End If

Case "Disc"
    Err.Clear
    On Error Resume Next
    ret = MsgBox("Disconnect between server and client?", vbYesNo + vbQuestion, "Disconnect")
    If ret = vbNo Then Exit Sub
    
    wsAccept.Close
    wsConnect.Close
    wsListen.Close
    
    If Err Then
        MsgBox "Error: " & Err.Description, vbCritical + vbOKOnly
    End If

Case "Save"
    'Record Data We Need
    Open App.Path & "\chatClient.txt" For Output As #1
    For i = 0 To lstChat.ListCount - 1
        Write #1, lstChat.List(i)
    Next i
    Close #1
    
    Open App.Path & "chatServer.txt" For Output As #1
    For i = 0 To lstServer.ListCount - 1
        Write #1, lstServer.List(i)
    Next i
    Close #1

Case "About"
mStr = "INTERNET DETECTIVE"
mStr = mStr + vbCrLf + vbCrLf + "Developed by Anoop.M.Nedumkunnam,anoopj13@yahoo.com"
mStr = mStr + vbCrLf + "Visit http://profiles.guru.com/anoopm"
MsgBox mStr, vbInformation, "About.."


End Select

End Sub

Private Sub wsAccept_Close()
wsConnect.Close

End Sub

Private Sub wsAccept_DataArrival(ByVal bytesTotal As Long)
wsAccept.GetData todat, vbString
lstChat.AddItem todat
On Error Resume Next
wsConnect.SendData todat
End Sub


Private Sub wsConnect_Close()
wsAccept.Close
YesConnect = True
End Sub

Private Sub wsConnect_Connect()
YesConnect = True
End Sub

Private Sub wsConnect_DataArrival(ByVal bytesTotal As Long)

wsConnect.GetData todat, vbString
lstServer.AddItem todat
On Error Resume Next
wsAccept.SendData todat
End Sub


Private Sub wsListen_ConnectionRequest(ByVal requestID As Long)

'Connects to the server
wsConnect.Connect gServerIp, gServerPort

On Error Resume Next

Do
'Just a looping task
dummy = DoEvents()
'We will make YesConnect true in the
'Connect event of wsConnect
If YesConnect = True Then wsAccept.Accept requestID
If YesConnect = True Then GoTo NoLoop
'b'coz i hate while statements..i don't know why. :)
Loop

NoLoop:

End Sub

