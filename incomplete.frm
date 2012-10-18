VERSION 5.00
Begin VB.Form incomplete 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Incomplete Feature"
   ClientHeight    =   1155
   ClientLeft      =   2475
   ClientTop       =   1755
   ClientWidth     =   2925
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "This feature is incomplete."
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "incomplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
