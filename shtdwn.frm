VERSION 5.00
Begin VB.Form ShutDownFRM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shut Down"
   ClientHeight    =   2115
   ClientLeft      =   2475
   ClientTop       =   2280
   ClientWidth     =   4470
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton shtdwnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton shtdwnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.ListBox shtdwnListBox 
      Height          =   840
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label shtdwnInfoText 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
End
Attribute VB_Name = "ShutDownFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
shtdwnListBox.AddItem "Shut Down"
shtdwnListBox.AddItem "Restart"
shtdwnListBox.AddItem "Log Off"

shtdwnInfoText.Caption = "To see what option does what, click it and this box will tell you."


End Sub

Private Sub shtdwnCancel_Click()
Unload Me
End Sub

Private Sub shtdwnListBox_Click()
Select Case shtdwnListBox.ListIndex
Case 0
 shtdwnInfoText.Caption = "Turns off the computer."
Case 1
 shtdwnInfoText.Caption = "Restarts the system, usually making it faster if a program crashed."
Case 2
 shtdwnInfoText.Caption = "Logs you out of the current session."
End Select
End Sub

Private Sub shtdwnOK_Click()
Select Case shtdwnListBox.ListIndex
Case 0
  
Case 1
 Shell ("shutdown /r")
Case 2
 Shell ("shutdown /l")
End Select
End Sub
