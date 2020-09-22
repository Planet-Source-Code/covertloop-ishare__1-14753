VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "iShare"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4680
      Top             =   6120
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00800000&
      Height          =   2260
      Left            =   2520
      ScaleHeight     =   2205
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   2520
      Width           =   3015
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   14
         TabIndex        =   0
         Top             =   1400
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Connect"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1710
         Width           =   1825
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Form1.frx":0CCA
         Left            =   1080
         List            =   "Form1.frx":0CD4
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   1815
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2000
         Left            =   0
         ScaleHeight     =   1995
         ScaleWidth      =   975
         TabIndex        =   4
         Top             =   210
         Width           =   975
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "v2.0"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   1680
            Width           =   975
         End
         Begin VB.Image Image3 
            Height          =   360
            Left            =   50
            Picture         =   "Form1.frx":0CF8
            Top             =   600
            Width           =   870
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   240
            Picture         =   "Form1.frx":1DBA
            Top             =   120
            Width           =   480
         End
         Begin VB.Image Image2 
            Height          =   570
            Left            =   180
            Picture         =   "Form1.frx":2A84
            Top             =   120
            Width           =   570
         End
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Connection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -120
         TabIndex        =   10
         Top             =   -15
         Width           =   3195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Nick Name:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1080
         TabIndex        =   8
         Top             =   1110
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Edition:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   1395
      End
   End
   Begin SHDocVwCtl.WebBrowser FileGet 
      Height          =   375
      Left            =   3360
      TabIndex        =   21
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
      ExtentX         =   661
      ExtentY         =   661
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSWinsockLib.Winsock ServerChatRelay 
      Index           =   0
      Left            =   2760
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock FileTransfer 
      Left            =   2760
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock FileTransferRequest 
      Left            =   2280
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ClientConnection 
      Left            =   2280
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Height          =   5295
      Left            =   0
      ScaleHeight     =   5235
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   0
      Width           =   855
      Begin VB.Image Image12 
         Height          =   480
         Left            =   150
         Picture         =   "Form1.frx":2E94
         ToolTipText     =   " Offline "
         Top             =   4560
         Width           =   480
      End
      Begin VB.Image Image9 
         Height          =   480
         Left            =   150
         MouseIcon       =   "Form1.frx":3B5E
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":3CB0
         ToolTipText     =   " Server Settings "
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   150
         MouseIcon       =   "Form1.frx":497A
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":4ACC
         ToolTipText     =   " Download "
         Top             =   3120
         Width           =   480
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   120
         MouseIcon       =   "Form1.frx":5796
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":58E8
         ToolTipText     =   " Server Status "
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   150
         MouseIcon       =   "Form1.frx":65B2
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":6704
         ToolTipText     =   " Close All Connections "
         Top             =   3840
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   150
         MouseIcon       =   "Form1.frx":73CE
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":7520
         ToolTipText     =   " Chat "
         Top             =   1680
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   150
         MouseIcon       =   "Form1.frx":7962
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":7AB4
         ToolTipText     =   " Connection Settings "
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Tools"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -120
         TabIndex        =   2
         Top             =   -15
         Width           =   1395
      End
   End
   Begin MSWinsockLib.Winsock ServerConnection 
      Left            =   2280
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00800000&
      Height          =   4515
      Left            =   840
      ScaleHeight     =   4455
      ScaleWidth      =   6915
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   6970
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Spit Server"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   295
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4080
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   3765
         ItemData        =   "Form1.frx":877E
         Left            =   4920
         List            =   "Form1.frx":8780
         TabIndex        =   16
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   0
         MaxLength       =   150
         TabIndex        =   14
         Top             =   4080
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Height          =   3735
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   6240
         TabIndex        =   23
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " X "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6680
         TabIndex        =   22
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "People Here:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   4920
         TabIndex        =   15
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Chat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -120
         TabIndex        =   12
         Top             =   -15
         Width           =   7500
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   5295
      Left            =   840
      ScaleHeight     =   5235
      ScaleWidth      =   6915
      TabIndex        =   17
      Top             =   -480
      Visible         =   0   'False
      Width           =   6970
      Begin SHDocVwCtl.WebBrowser PictureGet 
         Height          =   5010
         Left            =   0
         TabIndex        =   18
         Top             =   225
         Width           =   6910
         ExtentX         =   12188
         ExtentY         =   8837
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " X "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6600
         TabIndex        =   20
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Retrieving Picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -120
         TabIndex        =   19
         Top             =   -15
         Width           =   7000
      End
   End
   Begin VB.Image OnlinePic 
      Height          =   480
      Left            =   1440
      Picture         =   "Form1.frx":8782
      Top             =   6000
      Width           =   480
   End
   Begin VB.Image OfflinePic 
      Height          =   480
      Left            =   960
      Picture         =   "Form1.frx":944C
      Top             =   6000
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public i As Long
Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As Any) As Long

Private Declare Sub ReleaseCapture Lib "user32" ()

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub AddUnique(StringToAdd As String, lst As ListBox)
lst.Text = StringToAdd
If lst.ListIndex = -1 Then
lst.AddItem StringToAdd
End If
End Sub

Private Sub ClientConnection_Close()
Image12.Picture = OfflinePic.Picture
Picture5.Visible = False
Picture4.Visible = False
Picture2.Visible = True
End Sub

Private Sub ClientConnection_Connect()
Picture1.Enabled = True
Picture2.Enabled = True
Picture2.Visible = False
Image12.Picture = OnlinePic.Picture
AddUnique Text1.Text, List1
ClientConnection.SendData Text1.Text
End Sub

Private Sub ClientConnection_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim strChat As String
ClientConnection.GetData strChat
Dim Search, Where
Search = ":"
Where = InStr(strChat, Search)
If Where Then
Text2.Text = Text2.Text & vbCrLf & strChat
Text2.SetFocus
Text2.SelStart = Len(Text2.Text)
Text3.SetFocus
Else
AddUnique strChat, List1
End If
End Sub


Private Sub ClientConnection_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim Answers
Answers = MsgBox("The requested server is not online.", vbInformation + vbOKOnly, "iShare")
ClientConnection.Close
Picture1.Enabled = True
Picture2.Enabled = True
Exit Sub
End Sub


Private Sub Combo1_Click()
On Error Resume Next
Text1.SetFocus
End Sub


Private Sub Command1_Click()
If (Combo1.Text = "Client Edition") And (Form2.Text1.Text = "000.000.000.000") Or (Form2.Text1.Text = "") Or (Form2.Text2.Text = "") Then
Dim Answer
Answer = MsgBox("Invalid connection settings...", vbInformation + vbOKOnly, "iShare")
Image4_Click
Exit Sub
End If
Picture1.Enabled = False
Picture2.Enabled = False
If Combo1.Text = "Server Edition" Then
Picture1.Enabled = True
Picture2.Enabled = True
ServerConnection.Close
ServerConnection.LocalPort = CLng(187)
ServerConnection.Listen
FileTransferRequest.Close
FileTransferRequest.LocalPort = CLng(80)
FileTransferRequest.Listen
Picture2.Visible = False
Image12.Picture = OnlinePic.Picture
Exit Sub
End If
On Error GoTo erh
If Combo1.Text = "Client Edition" Then
ClientConnection.Close
ClientConnection.Connect Form2.Text1.Text, Form2.Text2.Text
Exit Sub
End If
erh:
Dim Answers
Answers = MsgBox("The requested server is not online.", vbInformation + vbOKOnly, "iShare")
ClientConnection.Close
Picture1.Enabled = True
Picture2.Enabled = True
Exit Sub
End Sub

Private Sub Command2_Click()
If Form4.Text1.Text = "" Then
Dim Answer
Answer = MsgBox("You have not set up your list." & Chr$(13) & "Would you like to do that now?", vbQuestion + vbYesNo, "iShare")
If Answer = 6 Then
Load Form4
Form4.Show
Exit Sub
End If
If Answer = 7 Then
Exit Sub
End If
Exit Sub
End If
Command2.Enabled = False
On Error Resume Next
Dim strChatSend As String
strChatSend = Text1.Text & ":  *** SERVER SPIT ***" & vbCrLf & vbCrLf & Form4.Text1.Text
For X = 1 To i
ServerChatRelay(X).SendData strChatSend
Pause (0.3)
Next X
Text2.Text = Text2.Text & vbCrLf & strChatSend
Text2.SetFocus
Text2.SelStart = Len(Text2.Text)
Text3.Text = ""
Text3.SetFocus
Command2.Enabled = True
End Sub

Private Sub FileGet_DownloadComplete()
Form1.MousePointer = 1
End Sub

Private Sub FileGet_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
Form1.MousePointer = 11
End Sub

Private Sub FileTransfer_DataArrival(ByVal bytesTotal As Long)
On Error GoTo erh
Form5.List1.AddItem "Request Admitted"
Dim strdata As String
FileTransfer.GetData strdata$ 'Get any data sent to us
If Mid$(strdata$, 1, 3) = "GET" Then 'If it is trying to get a site, find out
findget = InStr(strdata$, "GET ")      ' the site they want then set requestedPage$
spc2 = InStr(findget + 5, strdata$, " ") ' to it.
pagetoget$ = Mid$(strdata$, findget + 5, spc2 - (findget + 4))
RequestedPage$ = pagetoget$
Form5.List1.AddItem "Requested File:  " & RequestedPage$
'MsgBox requestedPage$  '<-- This is the file they are wanting
Form5.List1.AddItem "Sending Request"
f = FreeFile
temp = ""
If Form4.Dir1.Path = "C:\" Then
Open "C:\" & RequestedPage$ For Binary As #f
temp = Input(FileLen("C:\" & RequestedPage$), #f)
Close #f
End If
If Form4.Dir1.Path <> "C:\" Then
Open Form4.Label1.Caption & "\" & RequestedPage$ For Binary As #f
temp = Input(FileLen(Form4.Label1.Caption & "\" & RequestedPage$), #f)
Close #f
End If
getimg = temp
FileTransfer.SendData getimg
End If
Exit Sub
erh:
FileTransfer.SendData "<html><b>iShare Error</b><br><br>You have entered an invalid address and/or filename.<br>Check the address and/or filename for corrections.<br>You may use FileList.txt to manually retrieve the file list.<br><a href=" & Chr$(34) & "http://" & CurrentIP(True) & "/FileList.txt" & Chr$(34) & ">Or Click Here To Automatically Retrieve the File List</a></html>"
Form5.List1.AddItem "Request Error"
Exit Sub
End Sub

Private Sub FileTransfer_SendComplete()
FileTransfer.Close
Form5.List1.AddItem "Transfer Complete."
Form5.List1.AddItem "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
End Sub


Private Sub FileTransferRequest_ConnectionRequest(ByVal requestID As Long)
FileTransfer.Close
FileTransfer.Accept requestID
End Sub

Private Sub Form_Load()
'~~~~~This disregards a second load of iShare~~~~~'
'On Error GoTo erh
'AppActivate "iShare"
'Unload Me
'End
'Exit Sub
'erh:
i = 0
Picture2.Top = 1560
Picture2.Left = 2400
Combo1.ListIndex = 0
Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
ServerConnection.Close
FileTransferRequest.Close
ClientConnection.Close
FileTransfer.Close
For X = 0 To i
ServerChatRelay(X).Close
Next X
Unload Form2
Unload Me
End
End Sub


Private Sub Form_Terminate()
On Error Resume Next
ServerConnection.Close
FileTransferRequest.Close
ClientConnection.Close
FileTransfer.Close
For X = 0 To i
ServerChatRelay(X).Close
Next X
Unload Form2
Unload Me
End
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
ServerConnection.Close
FileTransferRequest.Close
ClientConnection.Close
FileTransfer.Close
For X = 0 To i
ServerChatRelay(X).Close
Next X
Unload Form2
Unload Me
End
End Sub


Private Sub Image4_Click()
Load Form2
If Combo1.Text = "Server Edition" Then Form2.SSTab1.Tab = 0
If Combo1.Text = "Client Edition" Then
Form2.SSTab1.Tab = 1
Form2.Command2.Enabled = True
End If
Form2.Show vbModal, Form1
End Sub

Private Sub Image5_Click()
If (Combo1.Text = "Client Edition") And (ClientConnection.State = 0) Then
Dim Answer
Answer = MsgBox("You are not connected.", vbInformation + vbOKOnly, "iShare")
Text1.SetFocus
Exit Sub
End If
If Combo1.Text = "Client Edition" Then Command2.Enabled = False
If Combo1.Text = "Server Edition" Then Command2.Enabled = True
Picture4.Visible = True
Text3.SetFocus
End Sub

Private Sub Image6_Click()
ServerConnection.Close
FileTransferRequest.Close
FileTransfer.Close
ClientConnection.Close
For X = 0 To i
ServerChatRelay(X).Close
Next X
Picture2.Top = 1560
Picture2.Left = 2400
Picture2.Visible = True
Picture5.Visible = False
Picture4.Visible = False
Image12.Picture = OfflinePic.Picture
End Sub


Private Sub Image7_Click()
If Combo1.Text = "Client Edition" Then Exit Sub
Load Form5
Form5.Show
End Sub

Private Sub Image8_Click()
If Combo1.Text = "Server Edition" Then Exit Sub
If (Combo1.Text = "Client Edition") And (ClientConnection.State = 0) Then
Dim Answer
Answer = MsgBox("You are not connected.", vbInformation + vbOKOnly, "iShare")
Text1.SetFocus
Exit Sub
End If
Load Form3
Form3.Show
End Sub

Private Sub Image9_Click()
If Combo1.Text = "Client Edition" Then Exit Sub
Load Form4
Form4.Show vbModal, Form1
End Sub


Private Sub Label1_DblClick()
Picture1.Top = 0
Picture1.Left = 0
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Picture1.hwnd, WM_NCLBUTTONDOWN, _
HTCAPTION, 0&)
End If
End Sub


Private Sub Label10_Click()
Picture4.Visible = False
End Sub

Private Sub Label5_DblClick()
Picture2.Top = 1560
Picture2.Left = 2400
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Picture2.hwnd, WM_NCLBUTTONDOWN, _
HTCAPTION, 0&)
End If
End Sub


Private Sub Label6_DblClick()
Picture4.Top = 0
Picture4.Left = 880
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Picture4.hwnd, WM_NCLBUTTONDOWN, _
HTCAPTION, 0&)
End If
End Sub


Private Sub Label8_DblClick()
Picture5.Left = 880
Picture5.Top = 0
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Picture5.hwnd, WM_NCLBUTTONDOWN, _
HTCAPTION, 0&)
End If
End Sub


Private Sub Label9_Click()
Picture5.Visible = False
End Sub


Private Sub ServerChatRelay_Close(Index As Integer)
Dim Uni
For X = 0 To List1.ListCount - 1
List1.ListIndex = X
If List1.Text = ServerChatRelay(Index).Tag Then
List1.RemoveItem List1.ListIndex
GoTo Uni
End If
Next X
Uni:
For J = 1 To i
For X = 0 To List1.ListCount - 1
On Error Resume Next
List1.ListIndex = X
ServerChatRelay(J).SendData List1.Text
Pause (0.3)
Next X
Pause (0.3)
Next J
Unload ServerChatRelay(Index)
End Sub

Private Sub ServerChatRelay_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim strChat As String
ServerChatRelay(Index).GetData strChat
Dim Search, Where
Search = ":"
Where = InStr(strChat, Search)
If Where Then
Text2.Text = Text2.Text & vbCrLf & strChat
Text2.SetFocus
Text2.SelStart = Len(Text2.Text)
For X = 1 To i
ServerChatRelay(X).SendData strChat
Pause (0.3)
Next X
Text3.SetFocus
Else
AddUnique strChat, List1
ServerChatRelay(Index).Tag = strChat
For J = 1 To i
For X = 0 To List1.ListCount - 1
List1.ListIndex = X
ServerChatRelay(J).SendData List1.Text
Pause (0.3)
Next X
Pause (0.3)
Next J
Exit Sub
End If
End Sub

Private Sub ServerConnection_ConnectionRequest(ByVal requestID As Long)
If i = 10 Then Exit Sub
i = i + 1
Load ServerChatRelay(i)
ServerChatRelay(i).Close
ServerChatRelay(i).Accept requestID
End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then Command1.Enabled = False
If Text1.Text <> "" Then Command1.Enabled = True
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
KeyAscii = 0
Exit Sub
End If
End Sub


Private Sub Web_StatusTextChange(ByVal Text As String)

End Sub


Private Sub Text2_GotFocus()
On Error Resume Next
Text3.SetFocus
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Dim strChatSend As String
strChatSend = Text1.Text & ":  " & Text3.Text
If Combo1.Text = "Server Edition" Then
For X = 1 To i
ServerChatRelay(X).SendData strChatSend
Pause (0.3)
Next X
Text2.Text = Text2.Text & vbCrLf & strChatSend
Text2.SetFocus
Text2.SelStart = Len(Text2.Text)
Text3.Text = ""
Text3.SetFocus
KeyAscii = 0
End If
If Combo1.Text = "Client Edition" Then
ClientConnection.SendData strChatSend
Text2.SetFocus
Text2.SelStart = Len(Text2.Text)
Text3.Text = ""
Text3.SetFocus
KeyAscii = 0
End If
End If
End Sub


Private Sub Timer1_Timer()
Label11.Caption = List1.ListCount
End Sub


