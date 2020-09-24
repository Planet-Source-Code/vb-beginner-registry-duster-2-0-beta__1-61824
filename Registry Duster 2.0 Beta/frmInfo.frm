VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Info"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2400
      Width           =   7335
   End
   Begin VB.TextBox txtSubKey 
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   7335
   End
   Begin VB.TextBox txtRootKey 
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   7335
   End
   Begin VB.TextBox txtFoundAt 
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      Caption         =   "Value"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   405
   End
   Begin VB.Label lblSubKey 
      AutoSize        =   -1  'True
      Caption         =   "Sub Key"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label lblRootKey 
      AutoSize        =   -1  'True
      Caption         =   "Root Key:"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   705
   End
   Begin VB.Label lblFoundAt 
      AutoSize        =   -1  'True
      Caption         =   "Found At:"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   690
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
txtRootKey.Text = frmMain.lvwRegErrors.SelectedItem.SubItems(1)
txtSubKey.Text = frmMain.lvwRegErrors.SelectedItem.SubItems(2)
txtValue.Text = frmMain.lvwRegErrors.SelectedItem.SubItems(3)
End Sub

