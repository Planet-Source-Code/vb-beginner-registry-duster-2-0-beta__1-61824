VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Registry Duster 2.0 Beta"
   ClientHeight    =   4035
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8985
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   8985
   Begin VB.Timer tmrScheduled 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ListView lvwRegErrors 
      Height          =   2175
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3836
      View            =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
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
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "&Start Scan"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8955
   End
   Begin VB.Label lblCurrentKey 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   7155
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblStatus 
      Caption         =   "Searching Key:"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1515
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuStartStop 
         Caption         =   "&Start Scan"
      End
      Begin VB.Menu mnuSeperator6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore Registry Backup"
      End
      Begin VB.Menu mnuSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRepair 
      Caption         =   "------------------>&Repair<------------------"
      Begin VB.Menu mnuCheckAll 
         Caption         =   "&Check All Items"
      End
      Begin VB.Menu mnuUncheckAll 
         Caption         =   "&Uncheck All Items"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFix 
         Caption         =   "&Delete All Checked Items"
      End
      Begin VB.Menu mnuSeperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search for Missing File, Manually (Experts Only)"
      End
      Begin VB.Menu mnuSeperator0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "&Info On This Item"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp2 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuSeperator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents cReg As cRegSearch
Attribute cReg.VB_VarHelpID = -1
Dim Recommended As Boolean

'Stop or start scanning for errors
Private Sub cmdStartStop_Click()
On Error GoTo ERROR_HANDLER

If cmdStartStop.Caption = "&Start Scan" Then mnuStartStop.Caption = "&Stop Scan"
If cmdStartStop.Caption = "&Stop Scan" Then
Caption = "Exiting..."
cReg.StopSearch
Exit Sub
End If

cmdStartStop.Caption = "&Stop Scan"
If lvwRegErrors.Visible = False Then
Top = Top / 2
Height = Height * 2
lvwRegErrors.Visible = True
End If
lvwRegErrors.ListItems.Clear
lblStatus.Caption = "Searching key:"
lblCurrentKey.Caption = ""

Load frmIgnore

'0=HKEY_ALL
cReg.RootKey = 0
'Don't search in any specific subkey (Search in all subkeys
cReg.SubKey = ""
'Only find errors in value names and value values
cReg.SearchFlags = KEY_NAME * 0 + VALUE_NAME * 1 + VALUE_VALUE * 1 + WHOLE_STRING * 0
'Search for registry values with the suffix "C:\"
cReg.SearchString = "C:\"
'Tell the user to wait
Me.Caption = "Scanning..."
'Start searching for invalid registry values
cReg.DoSearch

'If there are no items in lvwRegErrors, there is no need to keep mnuRepair visible
If lvwRegErrors.ListItems.Count = 0 Then
mnuRepair.Visible = False
Else
mnuRepair.Visible = True
End If

Exit Sub
ERROR_HANDLER:
End Sub

'The search is finished
Private Sub cReg_SearchFinished(ByVal lReason As Long)
If lReason = 0 Then
lblCurrentKey.Caption = "Done!"
ElseIf lReason = 1 Then
lblCurrentKey.Caption = "Terminated by user!"
Else
lblCurrentKey.Caption = "An Error occured! Err number = " & lReason
'Err.Raise lReason
End If
cmdStartStop.Caption = "&Start Scan"
mnuRepair.Visible = True
lblStatus.Caption = "Search result:"
Me.Caption = "Finished Scanning (" & lvwRegErrors.ListItems.Count & " errors found)"
End Sub

'If a registry error is found, add it to the list
Private Sub cReg_SearchFound(ByVal sRootKey As String, ByVal sKey As String, ByVal sValue As Variant, ByVal lFound As FOUND_WHERE)
Dim sTemp As String
Dim FileorPath As String
Dim lvItm As ListItem
Dim i As Integer
Dim ThePath As String

On Error GoTo ERROR_HANDLER

Select Case lFound
Case FOUND_IN_KEY_NAME
sTemp = "KEY_NAME"
Case FOUND_IN_VALUE_NAME
sTemp = "VALUE NAME"
Case FOUND_IN_VALUE_VALUE
sTemp = "VALUE VALUE"
End Select

FileorPath = sValue

'Check if the value is an invalid path or file
'If it is, then it adds the value to lvwRegErrors and displays the current number of errors, so far.
If FileorFolderExists(FormatValue(FileorPath)) = False Then

For i = 0 To frmIgnore.lstIgnore.ListCount - 1
If InStr(1, UCase$(FileorPath), frmIgnore.lstIgnore.List(i)) > 0 Or InStr(1, UCase$(sKey), frmIgnore.lstIgnore.List(i)) > 0 Then: Exit Sub
Next

If Recommended = True Then
ThePath = FormatValue(FileorPath)
If Right$(ThePath, 4) = ".EXE" Or Right$(ThePath, 4) = ".DLL" Or Right$(ThePath, 4) = ".OCX" Or Right$(ThePath, 4) = ".OCA" Or Right$(ThePath, 4) = ".SYS" Or Right$(ThePath, 4) = ".VXD" Or Right$(ThePath, 3) = ".AX" Or Right$(ThePath, 4) = ".COM" Or Right$(ThePath, 4) = ".CPL" Or Right$(ThePath, 4) = ".TLB" Or Right$(ThePath, 4) = ".SCR" Or Right$(ThePath, 4) = ".BAT" Or Right$(ThePath, 4) = ".DRV" Then
Else
Exit Sub
End If

End If

With lvwRegErrors
Set lvItm = .ListItems.Add(, , sTemp)
lvItm.SubItems(1) = sRootKey
lvItm.SubItems(2) = sKey
lvItm.SubItems(3) = sValue
End With
LV_AutoSizeColumn lvwRegErrors
Me.Caption = "Registry Duster 2.0 Beta (" & lvwRegErrors.ListItems.Count & " errors found)"
lblStatus.Caption = "Searching Key:" & vbCrLf & "(" & lvwRegErrors.ListItems.Count & " errors found)"

End If

Set lvItm = Nothing

Exit Sub
ERROR_HANDLER:
End Sub

'I don't know if I should remove this sub
Private Sub cReg_SearchKeyChanged(ByVal sFullKeyName As String)
'Note: This event cause a lot of printing.
'To increase performance remove this event.
Static tmr As Double
If Timer > tmr Then
tmr = Timer + 0.1
If Me.WindowState <> vbMinimized Then lblCurrentKey.Caption = sFullKeyName
End If
End Sub

'Setup everything
Private Sub Form_Load()
On Error GoTo ERROR_HANDLER

Set cReg = New cRegSearch

mnuRepair.Visible = False
With lvwRegErrors
.View = lvwReport
.ColumnHeaders.Add , , "Found at:"
.ColumnHeaders.Add , , "RootKey"
.ColumnHeaders.Add , , "SubKey"
.ColumnHeaders.Add , , "Value"
End With

Me.Move Me.Left, Me.Top, Screen.Width / 1.5, Screen.Height / 1.5
Me.Move (Screen.Width / 2) - (Me.ScaleWidth / 2), (Screen.Height / 2) - (Me.ScaleHeight / 2)

If GetSetting("Registry Duster", "Registry Duster", "Recommended") = 1 Then Recommended = True

If GetSetting("Registry Duster", "Registry Duster", "System Tray") = 1 Then AddToTray Me.Icon, Me.Caption, Me

Me.Top = GetSetting("Registry Duster", "Registry Duster", "Top")
Me.Left = GetSetting("Registry Duster", "Registry Duster", "Left")
Me.Width = GetSetting("Registry Duster", "Registry Duster", "Width")
Me.Height = GetSetting("Registry Duster", "Registry Duster", "Height")

Exit Sub
ERROR_HANDLER: SaveSetting "Registry Duster", "Registry Duster", "Recommended", 1
End Sub

'Respond to mouse events when this program is in the system tray
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ERROR_HANDLER
If RespondToTray(x) <> 0 Then PopupMenu mnuFile
Exit Sub
ERROR_HANDLER:
End Sub

'Save the window's position and size
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ERROR_HANDLER
SaveSetting "Registry Duster", "Registry Duster", "Top", Me.Top
SaveSetting "Registry Duster", "Registry Duster", "Left", Me.Left
SaveSetting "Registry Duster", "Registry Duster", "Width", Me.Width
SaveSetting "Registry Duster", "Registry Duster", "Height", Me.Height
Exit Sub
ERROR_HANDLER:
End Sub

'Resize the controls if the form is resized
Private Sub Form_Resize()
On Error GoTo ERROR_HANDLER
If Me.WindowState = vbMinimized Then Exit Sub
If Me.WindowState = vbMaximized Then lvwRegErrors.Visible = True
cmdStartStop.Move 0, cmdStartStop.Top, Me.ScaleWidth
cmdStartStop.Left = Me.ScaleWidth - cmdStartStop.Width
lblCurrentKey.Width = cmdStartStop.Left + cmdStartStop.Width - lblCurrentKey.Left
lvwRegErrors.Move 0, lblCurrentKey.Top + lblCurrentKey.Height, Me.ScaleWidth, Me.ScaleHeight - 1800
lvwRegErrors.ColumnHeaders(3).Width = (lvwRegErrors.Width - lvwRegErrors.ColumnHeaders(1).Width * 2) / 2 - 600
lvwRegErrors.ColumnHeaders(4).Width = lvwRegErrors.ColumnHeaders(3).Width
LV_AutoSizeColumn lvwRegErrors
Exit Sub
ERROR_HANDLER:
End Sub

'I'm not sure if this is necessary, but I guess it's just to clean up and exit this program
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ERROR_HANDLER
cReg.StopSearch
Set cReg = Nothing
End
Exit Sub
ERROR_HANDLER:
End Sub

'If the user clicks on a column, all the items will be selected
Private Sub lvwRegErrors_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If lvwRegErrors.ListItems.Count = 0 Then Exit Sub
mnuCheckAll_Click
Exit Sub
End Sub

'If you select multiple items, they will be checked if their unchecked and unchecked if their checked
Private Sub lvwRegErrors_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Item.Checked = True Then
Item.Checked = False
Else
Item.Checked = True
End If
End Sub

'If the the right clicks on lvwRegErrors, mnuRepair become visible
Private Sub lvwRegErrors_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 And mnuRepair.Visible = True Then
PopupMenu mnuRepair
End If
End Sub

'About menu
Private Sub mnuAbout_Click()
MsgBox "Registry Duster 2.0 Beta" & vbCrLf & vbCrLf & "If you use this program, you take full responsibility for any damages this program may do to your computer.", vbInformation, "Registry Duster 2.0 Beta"
End Sub

'Exit Registry Duster
Private Sub mnuExit_Click()
End
End Sub

'Creates a registry backup of all the values about to be deleted and then deletes them
Private Sub mnuFix_Click()
Dim i As Integer, nLoop As Single, m As Single
Dim Removed As Integer
Dim BackupFilename As String
On Error GoTo ERROR_HANDLER

'I don't think this is necessary, but if the registry backup takes a while, this program tells the user to wait.
lblCurrentKey.FontSize = 24
lblCurrentKey.FontBold = True
lblCurrentKey.Caption = "Creating Registry Backup..."
BackupReg
lblCurrentKey.FontSize = 8
lblCurrentKey.FontBold = False
lblCurrentKey.Caption = ""

BackupFilename = App.Path & "\RegBackup\Backup (" & Replace(Replace(Now, "/", "-"), ":", ";") & ").reg"

'Tell the user that this program has created a backup and and to restore the registry if the user's computer acts abnormal
MsgBox "This program has created a backup of all of the registry values that are about to be deleted. If you experience problems after using this, keep pressing F8 when you start up your computer and select Safe Mode and open up " & BackupFilename, vbInformation, "Important"

'Loop through every item in lvwRegErrors
For i = 1 To lvwRegErrors.ListItems.Count
'If the item is checked
If lvwRegErrors.ListItems.Item(i).Checked = True Then
'Delete the registry error and mark the item as removed
DeleteValue GetClassKey(lvwRegErrors.ListItems.Item(i).SubItems(1)), lvwRegErrors.ListItems.Item(i).SubItems(2), lvwRegErrors.ListItems.Item(i).SubItems(3)
lvwRegErrors.ListItems.Item(i).Text = "REMOVED"
lvwRegErrors.ListItems.Item(i).ForeColor = vbRed
Removed = Removed + 1
Else
lvwRegErrors.ListItems.Item(i).ForeColor = vbBlue
End If 'If you remove the if...then line above then also remove this line.
Next

'Tell the user how many items that were not removed

MsgBox "VB Registry Fixer has successfully fixed your registry. There were " & lvwRegErrors.ListItems.Count - Removed & " registry values that were NOT removed."

Exit Sub
ERROR_HANDLER:
End Sub

'Help menu
Private Sub mnuHelp2_Click()
MsgBox "Step 1 - Click Start Scan" & vbCrLf & vbCrLf & _
   "Step 2 - When the scan is finished, check all the items on the list that you want to delete. I highly recommend that you look carefully for what items you want to remove and not just check all of them." & vbCrLf & vbCrLf & _
   "Step 3 - Right click the list and click 'Delete All Checked Items'", vbInformation, "Help"
End Sub

'Checked all items in lvwRegErrors
Private Sub mnuCheckAll_Click()
Dim i As Integer
For i = 1 To lvwRegErrors.ListItems.Count
lvwRegErrors.ListItems.Item(i).Checked = True
Next
End Sub

Private Sub mnuInfo_Click()
frmInfo.Show
End Sub

'Show frmOptions
'For changing Registry Duster's settings
Private Sub mnuOptions_Click()
frmOptions.Show vbModal
End Sub

'Show frmRestore
'For restoring registry backups
Private Sub mnuRestore_Click()
frmRestore.Show vbModal
End Sub

'Shows frmSearch
'For searching for the file manually
Private Sub mnuSearch_Click()
Dim FileName As String
FileName = ReverseString(FormatValue(frmMain.lvwRegErrors.SelectedItem.SubItems(3)))
If Left$(FileName, 1) = "\" Then
FileName = Mid(FileName, 2)
End If
FileName = Mid(FileName, 1, InStr(1, FileName, "\") - 1)
FileName = ReverseString(FileName)
frmSearch.txtFile.Text = FileName
frmSearch.Show vbModal
End Sub

'Uncheck all checked items in lvwRegErrors
Private Sub mnuUncheckAll_Click()
Dim i As Integer
For i = 1 To lvwRegErrors.ListItems.Count
lvwRegErrors.ListItems.Item(i).Checked = False
Next
End Sub

'Start or stop the scan
Private Sub mnuStartStop_Click()
If mnuStartStop.Caption = "&Start Scan" Then
cmdStartStop_Click
mnuStartStop.Caption = "&Stop Scan"
End If

If mnuStartStop.Caption = "&Stop Scan" Then
cmdStartStop_Click
mnuStartStop.Caption = "&Start Scan"
End If
End Sub

'For scheduled scans
Private Sub tmrScheduled_Timer()
Dim TheDate As String
Dim Scheduled As String
Dim FormattedTime As String

On Error Resume Next

TheDate = Format(Now, "dddddd")
Scheduled = GetSetting("Registry Duster", "Registry Duster", "Scheduled")
If Mid(Scheduled, 1, 1) = 1 Then

FormattedTime = Replace(Time, Format(Time, "AM/PM"), "")

'If the user chooses to scan every hour
If Mid(Scheduled, 3) = "Hour" Then
If Val(Minute(FormattedTime)) = 0 Then
cmdStartStop_Click
End If
End If

'If the user chooses to scan every day
If Mid(Scheduled, 3) = "Day" Then
If Hour(FormattedTime) = GetSetting("Registry Duster", "Registry Duster", "Hour") Then
If Minute(FormattedTime) = GetSetting("Registry Duster", "Registry Duster", "Minute") Then
If Format(Time, "AM/PM") = GetSetting("Registry Duster", "Registry Duster", "Period") Then
cmdStartStop_Click
End If
End If
End If
End If

'If the user chooses to scan every week
If Mid(Scheduled, 3) = "Week" Then
'Day
If Format(Now, "dddd") = GetSetting("Registry Duster", "Registry Duster", "Day") Then
If Hour(FormattedTime) = GetSetting("Registry Duster", "Registry Duster", "Hour") Then
If Minute(FormattedTime) = GetSetting("Registry Duster", "Registry Duster", "Minute") Then
'AM or PM
If Format(Time, "AM/PM") = GetSetting("Registry Duster", "Registry Duster", "Period") Then
cmdStartStop_Click
End If
End If
End If
End If
End If

'If the user chooses to scan every month
If Mid(Scheduled, 3) = "Month" Then
If Format(Now, "mmmm") = GetSetting("Registry Duster", "Registry Duster", "Month") Then
'Hour
If Hour(FormattedTime) = "12" Then
'Minute
If Val(Minute(FormattedTime)) = 0 Then
'AM or PM
If Format(Time, "AM/PM") = "AM" Then
cmdStartStop_Click
End If
End If
End If
End If
End If
End If
End Sub
