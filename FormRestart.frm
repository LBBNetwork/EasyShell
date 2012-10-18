VERSION 5.00
Begin VB.Form FormRestart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FormRestart"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2085
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   2085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Changes saved, restart to see them"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FormRestart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FormRestart.Hide
End Sub


