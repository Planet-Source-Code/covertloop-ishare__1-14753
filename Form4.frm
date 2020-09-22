VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server File Share Settings"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   " DblClick to view complete directory name "
      Top             =   2280
      Width           =   2895
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function StrFormatByteSize Lib _
    "shlwapi" Alias "StrFormatByteSizeA" (ByVal _
    dw As Long, ByVal pszBuf As String, ByRef _
    cchBuf As Long) As String

Public Function FormatKB(ByVal Amount As Long) _
    As String
    Dim Buffer As String
    Dim Result As String
    Buffer = Space$(255)
    Result = StrFormatByteSize(Amount, Buffer, _
        Len(Buffer))
    If InStr(Result, vbNullChar) > 1 Then
        FormatKB = Left$(Result, InStr(Result, _
            vbNullChar) - 1)
    End If
End Function
Private Sub Command1_Click()
h = InputBox("Enter A File List Heading", "File List Heading")
If h = "" Then Text1.Text = Text1.Text & "***FILE LIST***"
If h <> "" Then Text1.Text = Text1.Text & h
For X = 0 To File1.ListCount - 1
File1.ListIndex = X
If Dir1.Path = "C:\" Then lFileSize = FileLen("C:\" & File1.FileName)
If Dir1.Path <> "C:\" Then lFileSize = FileLen(Label1.Caption & "\" & File1.FileName)
'MsgBox FormatKB(lFileSize)
Text1.Text = Text1.Text & vbCrLf & File1.FileName & " - FileSize:  " & FormatKB(lFileSize)
Next X
If Dir1.Path = "C:\" Then X = "C:\FileList.txt"
If Dir1.Path <> "C:\" Then X = Label1.Caption & "\FileList.txt"
Open X For Output As #1
Print #1, Text1.Text
Close #1
Form4.Hide
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Dir1_Change()
File1.Path = Dir1.Path
Dim sFile As String, sShortFile As String * 67
Dim lRet As Long
sFile = Dir1.Path
lRet = GetShortPathName(sFile, sShortFile, Len(sShortFile))
sFile = Left(sShortFile, lRet)
Label1.Caption = sFile
End Sub


Private Sub Drive1_Change()
On Error GoTo DriveHandler
Dir1.Path = Drive1.Drive
Exit Sub
DriveHandler:
Drive1.Drive = Dir1.Path
Exit Sub
End Sub


Private Sub Form_Load()
Dim sFile As String, sShortFile As String * 67
Dim lRet As Long
sFile = Dir1.Path
lRet = GetShortPathName(sFile, sShortFile, Len(sShortFile))
sFile = Left(sShortFile, lRet)
Label1.Caption = sFile
End Sub


Private Sub Label1_DblClick()
MsgBox Label1.Caption
End Sub


