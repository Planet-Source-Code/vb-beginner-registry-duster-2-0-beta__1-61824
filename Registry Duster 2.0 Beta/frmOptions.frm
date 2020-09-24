VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "Ignore List"
      Height          =   495
      Left            =   360
      TabIndex        =   20
      Top             =   4200
      Width           =   5415
   End
   Begin VB.Frame fmeFrame 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   5775
      Begin VB.OptionButton optMonth 
         Caption         =   "Month"
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtHour 
         Height          =   285
         Left            =   840
         MaxLength       =   2
         TabIndex        =   13
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton optHour 
         Caption         =   "Hour"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optDay 
         Caption         =   "Day at:"
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtMinute 
         Height          =   285
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   12
         Top             =   480
         Width           =   375
      End
      Begin VB.ComboBox cboPeriod 
         Height          =   315
         ItemData        =   "frmOptions.frx":0000
         Left            =   1920
         List            =   "frmOptions.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "Week on:"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cboDay 
         Height          =   315
         ItemData        =   "frmOptions.frx":0016
         Left            =   1080
         List            =   "frmOptions.frx":002F
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtHour2 
         Height          =   285
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   8
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtMinute2 
         Height          =   285
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   7
         Top             =   960
         Width           =   375
      End
      Begin VB.ComboBox cboPeriod2 
         Height          =   315
         ItemData        =   "frmOptions.frx":0073
         Left            =   3720
         List            =   "frmOptions.frx":007D
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblColon 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblAt 
         AutoSize        =   -1  'True
         Caption         =   "At:"
         Height          =   195
         Left            =   2400
         TabIndex        =   17
         Top             =   960
         Width           =   195
      End
      Begin VB.Label lblColon2 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3120
         TabIndex        =   16
         Top             =   960
         Width           =   75
      End
   End
   Begin VB.CheckBox chkScheduled 
      Caption         =   "Perform a scheduled scan every:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CheckBox chkStartup 
      Caption         =   "Run Registry Duster on Windows startup"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CheckBox chkSysTray 
      Caption         =   "Start Registry Duster minimized into the system tray"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CheckBox chkRecommended 
      Caption         =   "Only scan for files with .exe,.dll,.ocx,.oca,.sys,.vxd,.ax,.com,.cpl,.tlb,.scr,.bat and .drv extensions (Recommended)"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   240
      Top             =   4080
      Width           =   5655
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const HKEY_LOCAL_MACHINE = &H80000002

Private Sub chkRecommended_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkScheduled_Click()
If chkScheduled.Value = 1 Then
fmeFrame.Visible = True
Else
fmeFrame.Visible = False
End If
cmdApply.Enabled = True
End Sub

Private Sub chkStartup_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkSysTray_Click()
cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()

SaveSetting "Registry Duster", "Registry Duster", "Recommended", chkRecommended.Value
SaveSetting "Registry Duster", "Registry Duster", "System Tray", chkSysTray.Value

If chkStartup.Value = 1 Then
SetRegValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Registry Duster", App.Path & "\" & App.EXEName & ".exe"
Else
DeleteValue2 HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Registry Duster"
End If

If chkScheduled.Value = 1 Then
If optDay.Value = True Then
If txtHour.Text = "" Then MsgBox "Please type the hour in which your scheduled scan will run.", vbExclamation, "": txtHour.SetFocus: Exit Sub
If txtMinute.Text = "" Then MsgBox "Please type the minute in which your scheduled scan will run.", vbExclamation, "": txtMinute.SetFocus: Exit Sub
End If

If optWeek.Value = True Then
If txtHour2.Text = "" Then MsgBox "Please type the hour in which your scheduled scan will run.", vbExclamation, "": txtHour2.SetFocus: Exit Sub
If txtMinute2.Text = "" Then MsgBox "Please type the minute in which your scheduled scan will run.", vbExclamation, "": txtMinute2.SetFocus: Exit Sub
End If
End If

If Val(txtHour.Text) > 12 Then MsgBox "Please enter a correct hour.": txtHour.SetFocus: txtHour.SelStart = 0: txtHour.SelLength = Len(txtHour.Text): Exit Sub
If Val(txtMinute.Text) > 60 Then MsgBox "Please enter a correct minute.": txtMinute.SetFocus: txtMinute.SelStart = 0: txtMinute.SelLength = Len(txtMinute.Text): Exit Sub
If Val(txtHour2.Text) > 12 Then MsgBox "Please enter a correct hour.": txtHour2.SetFocus: txtHour2.SelStart = 0: txtHour2.SelLength = Len(txtHour2.Text): Exit Sub
If Val(txtMinute2.Text) > 60 Then MsgBox "Please enter a correct minute.": txtMinute2.SetFocus: txtMinute2.SelStart = 0: txtMinute2.SelLength = Len(txtMinute2.Text): Exit Sub

If chkStartup.Value = 0 And chkScheduled.Value = 1 Then MsgBox "In order to have scheduled scans, Registry Duster has to run when Windows starts.(The " & Chr(34) & "Start Registry Duster on Windows Startup" & Chr(34) & " checkbox will be checked)", vbExclamation, "": chkStartup.Value = 1: Exit Sub

If chkScheduled.Value = 1 Then
If optHour.Value = True Then
SaveSetting "Registry Duster", "Registry Duster", "Scheduled", chkScheduled.Value & "-Hour"
End If
If optDay.Value = True Then
SaveSetting "Registry Duster", "Registry Duster", "Scheduled", chkScheduled.Value & "-Day"
SaveSetting "Registry Duster", "Registry Duster", "Hour", txtHour.Text
SaveSetting "Registry Duster", "Registry Duster", "Minute", txtMinute.Text
SaveSetting "Registry Duster", "Registry Duster", "Period", cboPeriod.Text
End If
If optWeek.Value = True Then
SaveSetting "Registry Duster", "Registry Duster", "Scheduled", chkScheduled.Value & "-Week"
SaveSetting "Registry Duster", "Registry Duster", "Day", cboDay.Text
SaveSetting "Registry Duster", "Registry Duster", "Hour", txtHour2.Text
SaveSetting "Registry Duster", "Registry Duster", "Minute", txtMinute2.Text
SaveSetting "Registry Duster", "Registry Duster", "Period", cboPeriod2.Text
End If

If optMonth.Value = True Then
SaveSetting "Registry Duster", "Registry Duster", "Scheduled", chkScheduled.Value & "-Month"
End If
Else

SaveSetting "Registry Duster", "Registry Duster", "Scheduled", chkScheduled.Value

End If

SaveSetting "Registry Duster", "Registry Duster", "Startup", chkStartup.Value

Unload Me
End Sub

Private Sub cmdIgnore_Click()
frmIgnore.Show vbModal
End Sub

Private Sub Form_Load()
Dim Scheduled As String

On Error Resume Next

Scheduled = GetSetting("Registry Duster", "Registry Duster", "Scheduled")
chkRecommended.Value = GetSetting("Registry Duster", "Registry Duster", "Recommended")
chkSysTray.Value = GetSetting("Registry Duster", "Registry Duster", "System Tray")
chkStartup.Value = GetSetting("Registry Duster", "Registry Duster", "Startup")
chkScheduled.Value = Val(Mid(GetSetting("Registry Duster", "Registry Duster", "Scheduled"), 1, 1))
If Mid(Scheduled, 3) = "Hour" Then optHour.Value = True
If Mid(Scheduled, 3) = "Day" Then optDay.Value = True
If Mid(Scheduled, 3) = "Week" Then optWeek.Value = True
If Mid(Scheduled, 3) = "month" Then optMonth.Value = True
cboPeriod.Text = "AM"
cboDay.Text = "Sunday"
cboPeriod2.Text = "AM"
End Sub

'Enable and disable some controls
Private Sub optDay_Click()
cmdApply.Enabled = True
txtHour.Enabled = True
txtMinute.Enabled = True
lblColon.Enabled = True
cboPeriod.Enabled = True
cboDay.Enabled = False
lblAt.Enabled = False
txtHour2.Enabled = False
lblColon2.Enabled = False
cboPeriod2.Enabled = False
End Sub

Private Sub optHour_Click()
cmdApply.Enabled = True
End Sub

'Enable and disable some controls
Private Sub optWeek_Click()
cmdApply.Enabled = True
cboDay.Enabled = True
lblAt.Enabled = True
txtHour2.Enabled = True
lblColon2.Enabled = True
txtMinute2.Enabled = True
cboPeriod2.Enabled = True
txtHour.Enabled = False
txtMinute.Enabled = False
lblColon.Enabled = False
cboPeriod.Enabled = False
End Sub

'Only allow numbers
Private Sub txtHour_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57 '<--- this is for numbers
Exit Sub
Case Else
KeyAscii = 0
End Select
End Sub

'Only allow numbers
Private Sub txtHour2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57 '<--- this is for numbers
Exit Sub
Case Else
KeyAscii = 0
End Select
End Sub

'Only allow numbers
Private Sub txtMinute_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57 '<--- this is for numbers
Exit Sub
Case Else
KeyAscii = 0
End Select
End Sub

'Only allow numbers
Private Sub txtMinute2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57 '<--- this is for numbers
Exit Sub
Case Else
KeyAscii = 0
End Select
End Sub
