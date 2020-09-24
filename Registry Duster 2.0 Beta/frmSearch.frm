VERSION 5.00
Begin VB.Form frmSearch 
   Caption         =   "Search for Missing File"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   6495
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.ListBox lstFiles 
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   6855
   End
   Begin VB.Label lblCurrentFolder 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFile 
      Caption         =   "File to Search For:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUse 
         Caption         =   "&Use This File Instead"
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StopSearch As String

Private Sub cmdSearch_Click()
On Error GoTo ERROR_HANDLER
Dim SearchPath As String, FindStr As String
Dim FileSize As Long
Dim NumFiles As Integer, NumDirs As Integer
If cmdSearch.Caption = "&Search" Then
cmdSearch.Caption = "&Stop"
lstFiles.Clear
SearchPath = "C:\"
FindStr = txtFile.Text
FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
MsgBox NumFiles & " Files found in " & NumDirs + 1 & " Directories"
MsgBox "Size of files found under " & SearchPath & " = " & Format(FileSize, "#,###,###,##0") & " Bytes"
Screen.MousePointer = vbDefault
Else
cmdSearch.Caption = "&Search"
End If
Exit Sub
ERROR_HANDLER:
End Sub

Private Sub Form_Load()
mnuSeperator.Visible = False
mnuUse.Visible = False
End Sub

Private Sub Form_Resize()
txtFile.Move 0, txtFile.Top, Me.ScaleWidth, txtFile.Height
lstFiles.Move 0, lstFiles.Top, Me.ScaleWidth, Me.ScaleHeight
lblCurrentFolder.Move cmdSearch.Width, lblCurrentFolder.Top, Me.ScaleWidth - cmdSearch.Width, lblCurrentFolder.Height
End Sub

Private Sub lstFiles_DblClick()
Dim FilePath As String
FilePath = ReverseString(lstFiles.List(lstFiles.ListIndex))
If InStr(1, FilePath, "\") > 0 Then FilePath = Mid(FilePath, 2)
FilePath = Mid(FilePath, InStr(1, FilePath, "\") + 1)
FilePath = ReverseString(FilePath)
Shell "explorer.exe " & FilePath, vbMaximizedFocus
End Sub

Private Sub lstFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And cmdSearch.Caption = "&Search" And lstFiles.ListCount > 0 Then
mnuSeperator.Visible = True
mnuUse.Visible = True
PopupMenu mnuFile
End If
End Sub

Private Sub mnuUse_Click()
Dim FilePath As String
Dim FileData As String
Dim FileNum As String

On Error GoTo ERROR_HANDLER

FileNum = FreeFile

Open lstFiles.List(lstFiles.ListIndex) For Binary Access Read As #FileNum
FileData = Space$(LOF(1))
Get #FileNum, , FileData
Close #FileNum

FileNum = FreeFile

Open FormatValue(frmMain.lvwRegErrors.SelectedItem.SubItems(3)) For Binary Access Write As #FileNum
FileData = Space$(LOF(1))
Get #FileNum, , FileData
Close #FileNum

Exit Sub
ERROR_HANDLER:
MsgBox "An error occured. Maybe the folder doesn't exist where the file should be copied."
End Sub
