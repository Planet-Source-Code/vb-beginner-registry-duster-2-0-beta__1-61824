VERSION 5.00
Begin VB.Form frmIgnore 
   Caption         =   "Registry Values to Ignore"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstIgnore 
      Height          =   3180
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   6255
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   6240
      Width           =   2775
   End
   Begin VB.TextBox txtString 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   7335
   End
   Begin VB.Label lblString 
      Caption         =   "String to Ignore:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "If any part of a registry value matches the text in this list, Registry Duster will ignore the value, even if it's invalid."
      Height          =   390
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6225
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove"
      End
      Begin VB.Menu mnuSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmIgnore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
Dim MyString As String
Dim i As Integer

If txtString.Text = "" Then MsgBox "Please enter a string to ignore.", vbExclamation, "": txtString.SetFocus: Exit Sub
If Len(txtString.Text) = 1 Then MsgBox "The string must be more than one letter.", vbExclamation, "": txtString.SetFocus: txtString.SelStart = Len(txtString.Text): Exit Sub

MyString = UCase$(txtString.Text)
For i = 0 To lstIgnore.ListCount - 1
If MyString = lstIgnore.List(i) Then MsgBox "This string has already been added.", vbExclamation, "": txtString.SetFocus: txtString.SelStart = 0: txtString.SelLength = Len(txtString.Text): Exit Sub
Next

lstIgnore.AddItem MyString
End Sub

Private Sub cmdApply_Click()
Dim FileNum As Integer
Dim i As Integer
FileNum = FreeFile
Open App.Path & "\Ignore.txt" For Output As #FileNum
For i = 0 To lstIgnore.ListCount - 1
Print #1, lstIgnore.List(i)
Next
Close #FileNum
Unload Me
End Sub

Private Sub cmdRemove_Click()
If lstIgnore.ListIndex > -1 Then lstIgnore.RemoveItem lstIgnore.ListIndex
End Sub

Private Sub Form_Load()
Dim File, NextLine As String
Dim FileNum As Integer
Dim i As Integer
' clear the List Box
lstIgnore.Clear
' you want to load to the list box
File = App.Path & "\Ignore.txt"
' the FreeFile function assign unique number to the Filenum variable,
' to avoid collision with other opened file
FileNum = FreeFile
If FileorFolderExists(App.Path & "\Ignore.txt") = False Then Open File For Output As #FileNum: Close #FileNum

Open File For Input As #FileNum
' do until the file reach to its end
Do Until EOF(FileNum)
' read one line from the file to the NextLine String
Line Input #FileNum, NextLine
' add the line to the List Box
lstIgnore.AddItem UCase$(NextLine)
Loop
' Close the file
Close #FileNum

FileNum = FreeFile

RemoveDuplicates lstIgnore

If lstIgnore.ListCount > 0 Then
Do Until i = lstIgnore.ListCount - 1
If lstIgnore.List(i) = "" Then lstIgnore.RemoveItem i: i = i + 1
i = i + 1
Loop
End If
End Sub

'Resize controls if window is resized
Private Sub Form_Resize()
txtString.Width = Me.ScaleWidth
lblInfo.Width = Me.ScaleWidth
lstIgnore.Width = Me.ScaleWidth
cmdAdd.Left = (Me.ScaleWidth / 2) - (cmdAdd.Width / 2)
cmdApply.Left = (Me.ScaleWidth / 2) - (cmdApply.Width / 2)
End Sub

Private Sub lstIgnore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Position As Integer
Position = (Y / 180)
If Position > lstIgnore.ListCount Or Position = 0 Then Exit Sub
lstIgnore.Selected(Position - 1) = True
End Sub

Private Sub lstIgnore_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuFile
End If
End Sub

Private Sub mnuAdd_Click()
cmdAdd_Click
End Sub

Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub mnuRemove_Click()
cmdRemove_Click
End Sub
