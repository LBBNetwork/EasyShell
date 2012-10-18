VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ButtonDefine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   2730
   ClientLeft      =   2655
   ClientTop       =   1590
   ClientWidth     =   5175
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5175
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Programs"
      Height          =   1935
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton Command2 
         Caption         =   "Find program"
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtpath 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ComboBox combobutton1 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Path to program:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Select Button:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Button Captions"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox combobutton 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Rename button to:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Select Button:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   2160
      Width           =   855
   End
End
Attribute VB_Name = "ButtonDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Open "common\button1.cap" For Output As #1
'Print #1, Text1.Text
'Close #1
Select Case Text1.Text
Case ""

Case Else
 Select Case combobutton.ListIndex
 Case 0
  Open "common\button1.cap" For Output As #1
  Print #1, Text1.Text
  Close #1
 Case 1
  Open "common\button2.cap" For Output As #2
  Print #2, Text1.Text
  Close #2
 Case 2
  Open "common\button3.cap" For Output As #3
  Print #3, Text1.Text
  Close #3
 Case 3
  Open "common\button4.cap" For Output As #4
  Print #4, Text1.Text
  Close #4
 Case 4
  Open "common\button5.cap" For Output As #5
  Print #5, Text1.Text
  Close #5
 Case 5
  Open "common\button6.cap" For Output As #6
  Print #6, Text1.Text
  Close #6
 Case 6
  Open "common\button7.cap" For Output As #7
  Print #7, Text1.Text
  Close #7
 Case 7
  Open "common\button8.cap" For Output As #8
  Print #8, Text1.Text
  Close #8
 End Select
End Select

Dim quote As String

Select Case txtpath.Text
Case ""

Case Else
 Select Case combobutton1.ListIndex
 Case 0
  Open "common\button1.cfg" For Output As #9
  Print #9, txtpath.Text
  Close #9
 Case 1
  Open "common\button2.cfg" For Output As #9
  Print #9, txtpath.Text
  Close #9
 Case 2
  Open "common\button3.cfg" For Output As #10
  Print #10, txtpath.Text
  Close #10
 Case 3
  Open "common\button4.cfg" For Output As #11
  Print #11, txtpath.Text
  Close #11
 Case 4
  Open "common\button5.cfg" For Output As #12
  Print #12, txtpath.Text
  Close #12
 Case 5
  Open "common\button6.cfg" For Output As #13
  Print #13, txtpath.Text
  Close #13
 Case 6
  Open "common\button7.cfg" For Output As #14
  Print #14, txtpath.Text
  Close #14
 Case 7
  Open "common\button8.cfg" For Output As #14
  Print #14, txtpath.Text
  Close #14
 End Select
End Select

Unload Form1
Form1.Show
ButtonDefine.Hide
End Sub


Private Sub Command2_Click()

'Dim sfile As String

'CommonDialog1.DialogTitle = "Find program"
'CommonDialog1.Filter = "All Files (*.*)|*.*"
'CommonDialog1.InitDir = "C:\"
'CommonDialog1.ShowOpen

'If CommonDialog1 = "" Then
 
'Else
' sfile = CommonDialog1.FileName
'End If

'txtpath.Text = sfile

'Select Case CommonDialog1
'Case ""
 
'Case Else
' txtpath.Text = CommonDialog1
'End Select

  Dim sFile As String


   ' If ActiveForm Is Nothing Then LoadNewDoc
    

    With CommonDialog1
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "Program files (*.exe)|*.exe"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    'ActiveForm.rtfText.LoadFile sFile
    'ActiveForm.Caption = sFile
     txtpath.Text = sFile

End Sub

Private Sub Form_Load()
combobutton.AddItem "Button 1"
combobutton.AddItem "Button 2"
combobutton.AddItem "Button 3"
combobutton.AddItem "Button 4"
combobutton.AddItem "Button 5"
combobutton.AddItem "Button 6"
combobutton.AddItem "Button 7"
combobutton.AddItem "Button 8"

combobutton.ListIndex = 0

combobutton1.AddItem "Button 1"
combobutton1.AddItem "Button 2"
combobutton1.AddItem "Button 3"
combobutton1.AddItem "Button 4"
combobutton1.AddItem "Button 5"
combobutton1.AddItem "Button 6"
combobutton1.AddItem "Button 7"
combobutton1.AddItem "Button 8"

combobutton1.ListIndex = 0


End Sub

