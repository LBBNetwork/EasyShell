VERSION 5.00
Begin VB.Form WinShell32Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log In"
   ClientHeight    =   960
   ClientLeft      =   1965
   ClientTop       =   2280
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   4560
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Log In"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "WinShell32Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub
