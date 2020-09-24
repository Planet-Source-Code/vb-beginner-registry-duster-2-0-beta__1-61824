Attribute VB_Name = "mdlMain"
Option Explicit

'Start file searcher
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Const MAX_PATH = 260
Public Const MAXDWORD = &HFFFF
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName As String * MAX_PATH
cAlternate As String * 14
End Type
'End file searcher

Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Public Const HKEY_ALL = &H0&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003

Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

'Check if a path or file exists
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

'For ListView AutoSize
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const LVM_FIRST = &H1000

'System tray
Public Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type

Global Const NIM_ADD = &H0
Global Const NIM_MODIFY = &H1
Global Const NIM_DELETE = &H2

Global Const WM_MOUSEMOVE = &H200

Global Const NIF_MESSAGE = &H1
Global Const NIF_ICON = &H2
Global Const NIF_TIP = &H4

Global Const WM_LBUTTONDBLCLK = &H203
Global Const WM_RBUTTONUP = &H205

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Global nid As NOTIFYICONDATA

Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Private Const REG_SZ = 1

'For adding Registry Duster to Windows startup
Public Function SetRegValue(ROOTKEYS As ROOT_KEYS, Path As String, sKey As String, NewValue As String) As Boolean
Dim hKey As Long
Dim KeyValType As Long
Dim KeyValSize As Long
Dim KeyVal As String
Dim tmpVal As String
Dim res As Long
Dim i As Integer
Dim X As Long
On Error GoTo ERROR_HANDLER
res = RegOpenKeyEx(ROOTKEYS, Path, 0, KEY_ALL_ACCESS, hKey)
If res <> 0 Then GoTo ERROR_HANDLER
tmpVal = String(1024, 0)
KeyValSize = 1024
res = RegQueryValueEx(hKey, sKey, 0, KeyValType, tmpVal, KeyValSize)
Select Case res
Case 2
KeyValType = REG_SZ
Case Is <> 0
GoTo ERROR_HANDLER
End Select
Select Case KeyValType
Case REG_SZ
tmpVal = NewValue
End Select
KeyValSize = Len(tmpVal)
res = RegSetValueEx(hKey, sKey, 0, KeyValType, tmpVal, KeyValSize)
If res <> 0 Then GoTo ERROR_HANDLER
SetRegValue = True
RegCloseKey hKey
Exit Function
ERROR_HANDLER:
SetRegValue = False
RegCloseKey hKey
End Function

'Remove Chr(0)
Public Function StripNulls(OriginalStr As String) As String
If (InStr(OriginalStr, Chr(0)) > 0) Then
OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
End If
StripNulls = OriginalStr
End Function

'Reverse a string
Public Function ReverseString(TheString As String) As String
Dim i As Integer
For i = 1 To Len(TheString)
ReverseString = ReverseString & Mid(Right$(TheString, i), 1, 1)
Next
End Function

'Create a backup of the registry, I would have used the "regedit.exe /e" command but it takes too long.
Public Function BackupReg()
Dim i As Integer
Dim TheKey As String
Dim TheValue As String
Dim DefaultValue As Boolean
Dim BackupFilename As String
Dim PreviousKey As String
Dim FileNum As Integer

If FileorFolderExists(App.Path & "\RegBackups") = False Then MkDir App.Path & "\RegBackups"

BackupFilename = App.Path & "\RegBackups\Backup (" & Replace(Replace(Now, "/", "-"), ":", ";") & ").reg"

FileNum = FreeFile

Open BackupFilename For Output As #FileNum
Print #FileNum, "REGEDIT4"
'Loops through all the checked items and saves the values into C:\Backup.reg
For i = 1 To frmMain.lvwRegErrors.ListItems.Count
If frmMain.lvwRegErrors.ListItems.Item(i).Checked = True Then
TheKey = ReverseString(frmMain.lvwRegErrors.ListItems.Item(i).SubItems(1) & "\" & frmMain.lvwRegErrors.ListItems.Item(i).SubItems(2))
'I'm not sure, but I think that if the value ends with a "\", then it's the default value for that key
If Left$(TheKey, 1) = "\" Then DefaultValue = True: TheKey = Mid(TheKey, 2)
TheValue = Replace(ReverseString(Mid(TheKey, 1, InStr(1, TheKey, "\") - 1)), "\", "\\")
TheValue = Chr(34) & Replace(TheValue, Chr(34), "\" & Chr(34)) & Chr(34)
TheKey = ReverseString(Mid(TheKey, InStr(1, TheKey, "\") + 1))
If DefaultValue = True Then TheValue = "@"
If PreviousKey <> TheKey Then Print #FileNum, vbCrLf & "[" & TheKey & "]"
Print #FileNum, TheValue & "=" & Chr(34) & Replace(Replace(frmMain.lvwRegErrors.ListItems.Item(i).SubItems(3), "\", "\\"), Chr(34), "\" & Chr(34)) & Chr(34)
PreviousKey = TheKey
End If
Next

Close #FileNum
End Function

'Automatically resize the columns so that no text gets cut off
Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column As ColumnHeader = Nothing)
Dim C As ColumnHeader
If Column Is Nothing Then
For Each C In LV.ColumnHeaders
SendMessage LV.hwnd, LVM_FIRST + 30, C.Index - 1, -1
Next
Else
SendMessage LV.hwnd, LVM_FIRST + 30, Column.Index - 1, -1
End If
LV.Refresh
End Sub

'For searching for a file manually
Public Function FindFilesAPI(Path As String, SearchStr As String, FileCount As Integer, DirCount As Integer)
Dim FileName As String
Dim DirName As String
Dim dirNames() As String
Dim nDir As Integer
Dim i As Integer
Dim hSearch As Long
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer

If Right(Path, 1) <> "\" Then Path = Path & "\"
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(Path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
Do While Cont
If frmSearch.cmdSearch.Caption = "&Search" Then Exit Function
DirName = StripNulls(WFD.cFileName)
If (DirName <> ".") And (DirName <> "..") Then
If GetFileAttributes(Path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
dirNames(nDir) = DirName
frmSearch.lblCurrentFolder.Refresh
frmSearch.lblCurrentFolder.Caption = Path & DirName
DirCount = DirCount + 1
nDir = nDir + 1
ReDim Preserve dirNames(nDir)
End If
End If
Cont = FindNextFile(hSearch, WFD)
DoEvents
Loop
Cont = FindClose(hSearch)
End If
hSearch = FindFirstFile(Path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
While Cont
If frmSearch.cmdSearch.Caption = "&Search" Then Exit Function
FileName = StripNulls(WFD.cFileName)
frmSearch.Refresh
frmSearch.lblCurrentFolder.Caption = Path
If (FileName <> ".") And (FileName <> "..") Then
FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
FileCount = FileCount + 1
frmSearch.lstFiles.AddItem Path & FileName
End If
Cont = FindNextFile(hSearch, WFD)
DoEvents
Wend
Cont = FindClose(hSearch)
End If
If nDir > 0 Then
For i = 0 To nDir - 1
If frmSearch.cmdSearch.Caption = "&Search" Then Exit Function
FindFilesAPI = FindFilesAPI + FindFilesAPI(Path & dirNames(i) & "\", SearchStr, FileCount, DirCount)
DoEvents
Next i
End If
End Function

'A DeleteValue function
Public Sub DeleteValue(ROOTKEYS As ROOT_KEYS, Path As String, sKey As String)
Dim ValKey As String
Dim SecKey As String, SlashPos As Single
SlashPos = InStrRev(Path, "\", compare:=vbTextCompare)
SecKey = Left(Path, SlashPos - 1) 'This will retreive the section key that I need
ValKey = Right(Path, Len(Path) - SlashPos) 'This will retreive the ValueKey that I need to delete
DeleteValue2 ROOTKEYS, SecKey, ValKey
End Sub

'Another DeleteValue function
Public Sub DeleteValue2(hKey As ROOT_KEYS, strPath As String, strValue As String)
Dim Ret
RegCreateKey hKey, strPath, Ret
RegDeleteValue Ret, strValue
RegCloseKey Ret
End Sub

'Returns the long value of the string entered.
Public Function GetClassKey(cls As String) As ROOT_KEYS
Select Case cls
Case "HKEY_ALL"
GetClassKey = HKEY_ALL
Case "HKEY_CLASSES_ROOT"
GetClassKey = HKEY_CLASSES_ROOT
Case "HKEY_CURRENT_USER"
GetClassKey = HKEY_CURRENT_USER
Case "HKEY_LOCAL_MACHINE"
GetClassKey = HKEY_LOCAL_MACHINE
Case "HKEY_USERS"
GetClassKey = HKEY_USERS
Case "HKEY_PERFORMANCE_DATA"
GetClassKey = HKEY_PERFORMANCE_DATA
Case "HKEY_CURRENT_CONFIG"
GetClassKey = HKEY_CURRENT_CONFIG
Case "HKEY_DYN_DATA"
GetClassKey = HKEY_DYN_DATA
End Select
End Function

'Checks if a folder or file exists
Public Function FileorFolderExists(FolderOrFilename As String) As Boolean
If PathFileExists(FolderOrFilename) = 1 Then
FileorFolderExists = True
ElseIf PathFileExists(FolderOrFilename) = 0 Then
FileorFolderExists = False
End If
End Function

'Format the values by removing everything except for the filename or path
Public Function FormatValue(sValue As String) As String
Dim FileorPath As String

'Fix up the file or path so that it's compatible with the FileorFolderExists function
FileorPath = UCase$(sValue)

'I don't know if it's just my computer, but some registry values somehow didn't contain "C:\"
If InStr(1, FileorPath, "C:\") = 0 Then FileorPath = "C:\"

'Find the start of the path or filename (Example:"h6j65ej(C:\Test)")
FileorPath = Mid(FileorPath, InStr(1, FileorPath, "C:\"))

'The same as the other one except this is for specific extensions
'(Example:"C:\lalalalalalalala\idfjb.dll\50")
'This won't work if you have a folder named something like "blabla.exe"
If InStr(1, FileorPath, ".EXE") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".EXE") + 3)
If InStr(1, FileorPath, ".SYS") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".SYS") + 3)
If InStr(1, FileorPath, ".DLL") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".DLL") + 3)
If InStr(1, FileorPath, ".OCX") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".OCX") + 3)
If InStr(1, FileorPath, ".OCA") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".OCA") + 3)
If InStr(1, FileorPath, ".VXD") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".VXD") + 3)
If InStr(1, FileorPath, ".AX") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".AX") + 2)
If InStr(1, FileorPath, ".COM") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".COM") + 3)
If InStr(1, FileorPath, ".CPL") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".CPL") + 3)
If InStr(1, FileorPath, ".TLB") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".TLB") + 3)
If InStr(1, FileorPath, ".SCR") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".SCR") + 3)
If InStr(1, FileorPath, ".BAT") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".BAT") + 3)
If InStr(1, FileorPath, ".DRV") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".DRV") + 3)

'Remove everything after the path. This definitely doesn't work for all values.
'(Example:"C:\blablablablablablablabla?5784846\84585")
If InStr(1, FileorPath, "/") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "/") - 1)
If InStr(1, FileorPath, "*") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "*") - 1)
If InStr(1, FileorPath, "?") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "?") - 1)
If InStr(1, FileorPath, Chr(34)) > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, Chr(34)) - 1)
If InStr(1, FileorPath, "<") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "<") - 1)
If InStr(1, FileorPath, ">") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ">") - 1)
If InStr(1, FileorPath, "|") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "|") - 1)
If InStr(1, FileorPath, ",") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ",") - 1)
If InStr(1, FileorPath, "(") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "(") - 1)
If InStr(1, FileorPath, ";") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ";") - 1)
If InStr(3, FileorPath, ":") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ":") - 1)


'%1 is used for file associations
'(Example:"C:\WINDOWS\NOTEPAD.EXE %1")
FormatValue = Replace(FileorPath, " %1", "")
End Function

'Add Registry Duster to the system tray
Sub AddToTray(TrayIcon, TrayText As String, TrayForm As Form)
nid.cbSize = Len(nid)
nid.hwnd = TrayForm.hwnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = TrayIcon
nid.szTip = TrayText & vbNullChar

Shell_NotifyIcon NIM_ADD, nid
TrayForm.Hide
End Sub

'Add Registry Duster to the system tray
Sub ModifyTray(TrayIcon, TrayText As String, TrayForm As Form)
nid.cbSize = Len(nid)
nid.hwnd = TrayForm.hwnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = TrayIcon
nid.szTip = TrayText & vbNullChar

Shell_NotifyIcon NIM_MODIFY, nid
End Sub

'Remove the Registry Duster icon from the system tray
Sub RemoveFromTray()
Shell_NotifyIcon NIM_DELETE, nid
End Sub

'Registry Duster responds to mouse events
Function RespondToTray(X As Single)
RespondToTray = 0
Dim msg As Long
Dim sFilter As String
If frmMain.ScaleMode <> 3 Then msg = X / Screen.TwipsPerPixelX Else: msg = X
Select Case msg
Case WM_LBUTTONDBLCLK
frmMain.WindowState = 0
frmMain.Show
Case WM_RBUTTONUP
frmMain.PopupMenu frmMain.mnuFile
End Select
End Function

'A function for removing duplicates in a ListBox
Public Function RemoveDuplicates(ListBox As ListBox)
Dim Col As New Collection
Dim i As Long
On Error Resume Next

If ListBox.ListCount > 1 Then

For i = 0 To ListBox.ListCount - 1
Col.Add ListBox.List(i), ListBox.List(i)
Next
ListBox.Clear

For i = 1 To Col.Count
ListBox.AddItem Col.Item(i)
Next
Set Col = Nothing
End If
End Function

'Log errors
Public Function LogError(Detail As String)
Dim FileNum As Integer
FileNum = FreeFile
Open App.Path & "\ErrorLog.txt" For Output As #1
Print #1, Now & vbCrLf & "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description & vbCrLf & "Detail: " & Detail & vbCrLf
Close #1
End Function
