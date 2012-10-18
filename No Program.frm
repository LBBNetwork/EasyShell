VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "No program to run"
   ClientHeight    =   1680
   ClientLeft      =   2835
   ClientTop       =   2280
   ClientWidth     =   3945
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Please contact your system administrator to add a program to this button."
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "There is no program assigned to this button."
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
