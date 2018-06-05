VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H000000FF&
   Caption         =   "Summative Zeldin"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   5
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmdComputerScienceQuiz 
      Caption         =   "Computer Science Quiz (T3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   4
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton cmdRemainderMethod 
      Caption         =   "Remainder Method"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdBooleanGates 
      Caption         =   "Boolean Logic Gates (T3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdGuessingGame 
      Caption         =   "Guessing Game (T3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      Caption         =   "ICS2OG Summative 2018: By Josh Z."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Zeldin
Option Explicit
Private Sub cmdReturn1_Click()
    Unload frmGuessingGame
End Sub

Private Sub cmdBooleanGates_Click()
    frmLogicGates.Show vbModal
End Sub

Private Sub cmdComputerScienceQuiz_Click()
    frmComputerScienceQuiz.Show vbModal
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdGuessingGame_Click()
    frmGuessingGame.Show vbModal
End Sub


Private Sub cmdRemainderMethod_Click()
    frmRemainderMethod.Show vbModal
End Sub
