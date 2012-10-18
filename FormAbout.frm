VERSION 5.00
Begin VB.Form FormAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EasyShell"
   ClientHeight    =   2745
   ClientLeft      =   2475
   ClientTop       =   2100
   ClientWidth     =   3765
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   120
      Picture         =   "FormAbout.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "http://www.beige-box.com"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "Program (c) 2010 The Little Beige Box"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label6 
      Caption         =   "Uses some code (c) 2010 Jecag "
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "Version 1.0"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   840
      X2              =   3720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "EasyShell"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

Label2.Caption = "Build " & App.Revision & ""
End Sub
