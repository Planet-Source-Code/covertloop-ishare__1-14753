VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Download"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2640
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   2640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get It"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter The Name Of The File To Retrieve:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2385
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'GIF Format
Dim Searchgif, Where
Searchgif = ".gif"
Where = InStr(Text1.Text, Searchgif)
If Where Then
Form1.Picture5.Left = 880
Form1.Picture5.Top = 0
Form1.Picture5.Visible = True
Form1.PictureGet.Navigate "http://" & Form1.ClientConnection.RemoteHostIP & "/" & Text1.Text
Unload Me
Exit Sub
End If
Dim Searchgiff, Wheree
Searchgiff = ".Gif"
Wheree = InStr(Text1.Text, Searchgiff)
If Wheree Then
Form1.Picture5.Left = 880
Form1.Picture5.Top = 0
Form1.Picture5.Visible = True
Form1.PictureGet.Navigate "http://" & Form1.ClientConnection.RemoteHostIP & "/" & Text1.Text
Unload Me
Exit Sub
End If
Dim Searchgifff, Whereee
Searchgifff = ".GIF"
Whereee = InStr(Text1.Text, Searchgifff)
If Whereee Then
Form1.Picture5.Left = 880
Form1.Picture5.Top = 0
Form1.Picture5.Visible = True
Form1.PictureGet.Navigate "http://" & Form1.ClientConnection.RemoteHostIP & "/" & Text1.Text
Unload Me
Exit Sub
End If

'JPG Format
Dim Searchjpg, Wherea
Searchjpg = ".jpg"
Wherea = InStr(Text1.Text, Searchjpg)
If Wherea Then
Form1.Picture5.Left = 880
Form1.Picture5.Top = 0
Form1.Picture5.Visible = True
Form1.PictureGet.Navigate "http://" & Form1.ClientConnection.RemoteHostIP & "/" & Text1.Text
Unload Me
Exit Sub
End If
Dim Searchjpgg, Whereaa
Searchjpgg = ".Jpg"
Whereaa = InStr(Text1.Text, Searchjpgg)
If Whereaa Then
Form1.Picture5.Left = 880
Form1.Picture5.Top = 0
Form1.Picture5.Visible = True
Form1.PictureGet.Navigate "http://" & Form1.ClientConnection.RemoteHostIP & "/" & Text1.Text
Unload Me
Exit Sub
End If
Dim Searchjpggg, Whereaaa
Searchjpggg = ".JPG"
Whereaaa = InStr(Text1.Text, Searchjpggg)
If Whereaaa Then
Form1.Picture5.Left = 880
Form1.Picture5.Top = 0
Form1.Picture5.Visible = True
Form1.PictureGet.Navigate "http://" & Form1.ClientConnection.RemoteHostIP & "/" & Text1.Text
Unload Me
Exit Sub
End If
Form1.FileGet.Navigate "http://" & Form1.ClientConnection.RemoteHostIP & "/" & Text1.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Form_Load()
Text1.SelStart = Len(Text1.Text)
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


