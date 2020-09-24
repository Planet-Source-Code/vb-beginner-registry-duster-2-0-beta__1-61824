VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRestore 
   Caption         =   "Restore Registry Backup(s)"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox fleBackups 
      Height          =   2235
      Left            =   0
      Pattern         =   "*.reg"
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComctlLib.ListView lvwRegBackups 
      Height          =   4215
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7435
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore Backup(s)"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Dim FileName As String
Dim i As Integer
Dim lvItm As ListItem

If FileorFolderExists(App.Path & "\RegBackups") = False Then MkDir App.Path & "\RegBackups"
fleBackups.Path = "RegBackups"
With lvwRegBackups
.View = lvwReport
.ColumnHeaders.Add , , "Backup Number"
.ColumnHeaders.Add , , "Date Created"
.ColumnHeaders.Add , , "Filesize (in bytes)"
End With

For i = 1 To fleBackups.ListCount
FileName = fleBackups.List(i - 1)
With lvwRegBackups
Set lvItm = .ListItems.Add(, , "Backup #" & i)
lvItm.SubItems(1) = Mid(Mid(Replace(Replace(FileName, "-", "/"), ";", ":"), 1, Len(FileName) - 5), InStr(1, FileName, "(") + 1)
lvItm.SubItems(2) = Len(FileName)
End With
lvwRegBackups.ListItems.Item(i).Tag = FileName
Next i

Me.Move (Screen.Width / 2) - (Me.ScaleWidth / 2), (Screen.Height / 2) - (Me.ScaleHeight / 2)
lvwRegBackups.ColumnHeaders(1).Width = lvwRegBackups.Width / 3
lvwRegBackups.ColumnHeaders(2).Width = lvwRegBackups.Width / 3
lvwRegBackups.ColumnHeaders(3).Width = lvwRegBackups.Width / 3

If lvwRegBackups.ListItems.Count = 0 Then MsgBox "Sorry, no backups found.": Unload Me

Set lvItm = Nothing

End Sub

Private Sub Form_Resize()
lvwRegBackups.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub lvwRegBackups_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Item.Checked = True Then
Item.Checked = False
Else
Item.Checked = True
End If
End Sub

Private Sub lvwRegBackups_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuFile
End If
End Sub

Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub mnuRestore_Click()
Dim i As Integer
Dim ItemsChecked As Integer

If FileorFolderExists(App.Path & "\RegBackups") = False Then MkDir App.Path & "\RegBackups"
For i = 1 To lvwRegBackups.ListItems.Count
If lvwRegBackups.ListItems.Item(i).Checked = True Then
MsgBox lvwRegBackups.ListItems.Item(i).Tag
Shell "regedit.exe /s" & Chr(34) & App.Path & "\RegBackups\" & lvwRegBackups.ListItems.Item(i).Tag & Chr(34)
MsgBox "The registry was be restored successfully."
ItemsChecked = ItemsChecked + 1
End If
Next

If ItemsChecked = 0 Then MsgBox "Please click the checkbox(es) next to the item you would like to restore.", vbExclamation, "None Checked"

End Sub
