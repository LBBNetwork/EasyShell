VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   1425
   ClientLeft      =   1920
   ClientTop       =   1905
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "EasyShell"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.CommandButton cmdShutdown 
         Caption         =   "Shut Down"
         Height          =   255
         Left            =   7440
         TabIndex        =   17
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton defineprog 
         Height          =   615
         Left            =   7800
         Picture         =   "Form1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Frame Frame4 
         Caption         =   "Your Programs"
         Height          =   975
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   5895
         Begin VB.CommandButton program8 
            Caption         =   "Program 8"
            Height          =   255
            Left            =   4440
            TabIndex        =   15
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton program7 
            Caption         =   "Program 7"
            Height          =   255
            Left            =   3000
            TabIndex        =   14
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton program6 
            Caption         =   "Program 6"
            Height          =   255
            Left            =   1560
            TabIndex        =   13
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton program5 
            Caption         =   "Program 5"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton program4 
            Caption         =   "Program 4"
            Height          =   255
            Left            =   4440
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton program3 
            Caption         =   "Program 3"
            Height          =   255
            Left            =   3000
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Program2 
            Caption         =   "Program 2"
            Height          =   255
            Left            =   1560
            TabIndex        =   9
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Program1 
            Caption         =   "Program 1"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton AboutShow 
         Appearance      =   0  'Flat
         Caption         =   "About"
         Height          =   255
         Left            =   240
         Picture         =   "Form1.frx":0442
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "Date"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
         Begin VB.Label LBLDate 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   9360
         Top             =   840
      End
      Begin VB.Frame Frame2 
         Caption         =   "Time"
         Height          =   615
         Left            =   9120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Line Line2 
         BorderStyle     =   6  'Inside Solid
         X1              =   9000
         X2              =   9000
         Y1              =   120
         Y2              =   1320
      End
      Begin VB.Line Line1 
         BorderStyle     =   6  'Inside Solid
         X1              =   1320
         X2              =   1320
         Y1              =   120
         Y2              =   1320
      End
      Begin VB.Label EvalVersion 
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         Height          =   255
         Left            =   9720
         TabIndex        =   1
         Top             =   1080
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub ShutDownDialog Lib "shell32.dll" Alias "#60" (ByVal hwndOwner As Long)

Private Sub AboutShow_Click()
FormAbout.Show
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub cmdShutdown_Click()
'ShutDownFRM.Show
ShutDownDialog hWnd

'shtdwn = MsgBox("To be fair to all users, the shutdown functionality is disabled.")
End Sub

Private Sub Command2_Click()
WinShell32Login.Show
End Sub

Private Sub defineprog_Click()
ButtonDefine.Show
End Sub


Private Sub Form_Load()

On Error GoTo errorhandle

Dim button1cap As String
Dim ProductName As String
ProductName = "EasyShell"

Frame1.Caption = ProductName
'EvalVersion.Caption = "Build " & App.Revision
Label2.Caption = Time
LBLDate.Caption = Date

Open "common\button1.cap" For Input As #1
Input #1, button1cap
Close #1

Open "common\button2.cap" For Input As #2
Input #2, button2cap
Close #2

Open "common\button3.cap" For Input As #3
Input #3, button3cap
Close #3

Open "common\button4.cap" For Input As #4
Input #4, button4cap
Close #4

Open "Common\button5.cap" For Input As #5
Input #5, button5cap
Close #5

Open "common\button6.cap" For Input As #6
Input #6, button6cap
Close #6

Open "common\button7.cap" For Input As #7
Input #7, button7cap
Close #7

Open "common\button8.cap" For Input As #8
Input #8, button8cap
Close #8

Program1.Caption = button1cap
Program2.Caption = button2cap
program3.Caption = button3cap
program4.Caption = button4cap
program5.Caption = button5cap
program6.Caption = button6cap
program7.Caption = button7cap
program8.Caption = button8cap

errorhandle:
Select Case Err
Case Is <> 0
 errdlg = MsgBox("Could not load the configuration files." & vbNewLine & "WinShell32 will now close.", vbCritical, "Critical Error")
 End
End Select
End Sub



Private Sub Label1_Click()
End
End Sub



Private Sub Program1_Click()
Open "common\button1.cfg" For Input As #2
Input #2, button1$
Close #2

Select Case button1$
Case "none"
 NoButton = MsgBox("Nothing is assigned to this button.", vbExclamation, "No program set")
Case Else
 On Error GoTo launcherr
  X% = Shell(button1$, 1)
End Select

launcherr:
'RunProg = MsgBox("Error running the program." & vbNewLine & "It may not exist in the specified path.", vbCritical, "Error loading program")


End Sub


Private Sub Program2_Click()
Open "common\button2.cfg" For Input As #9
Input #9, button2$
Close #9
Select Case button2$
Case "none"
 'Form3.Show
  NoButton = MsgBox("Nothing is assigned to this button.", vbExclamation, "No program set")
Case Else
 X% = Shell(button2$, 1)
End Select
End Sub

Private Sub program3_Click()
Open "common\button3.cfg" For Input As #10
Input #10, button3$
Close #10
Select Case button3$
Case "none"
 'Form3.Show
  NoButton = MsgBox("Nothing is assigned to this button.", vbExclamation, "No program set")
Case Else
 X% = Shell(button3$, 1)
End Select

End Sub

Private Sub program4_Click()
Open "common\button4.cfg" For Input As #11
Input #11, button4$
Close #11
Select Case button4$
Case "none"
 'Form3.Show
  NoButton = MsgBox("Nothing is assigned to this button.", vbExclamation, "No program set")
Case Else
 X% = Shell(button4$, 1)
End Select
End Sub

Private Sub program5_Click()
Open "common\button5.cfg" For Input As #12
Input #12, button5$
Close #12
Select Case button5$
Case "none"
 'Form3.Show
  NoButton = MsgBox("Nothing is assigned to this button.", vbExclamation, "No program set")
Case Else
 X% = Shell(button5$, 1)
End Select
End Sub

Private Sub program6_Click()
Open "common\button6.cfg" For Input As #13
Input #13, button6$
Close #13
Select Case button6$
Case "none"
 'Form3.Show
  NoButton = MsgBox("Nothing is assigned to this button.", vbExclamation, "No program set")
Case Else
 X% = Shell(button6$, 1)
End Select
End Sub

Private Sub program7_Click()
Open "common\button7.cfg" For Input As #14
Input #14, button7$
Close #14
Select Case button7$
Case "none"
 'Form3.Show
  NoButton = MsgBox("Nothing is assigned to this button.", vbExclamation, "No program set")
Case Else
 X% = Shell(button7$, 1)
End Select
End Sub

Private Sub program8_Click()
Open "common\button8.cfg" For Input As #15
Input #15, button8$
Close #15
Select Case button8$
Case "none"
 'Form3.Show
  NoButton = MsgBox("Nothing is assigned to this button.", vbExclamation, "No program set")
Case Else
 X% = Shell(button8$, 1)
End Select
End Sub

Private Sub Timer1_Timer()
Label2.Caption = Time
End Sub
