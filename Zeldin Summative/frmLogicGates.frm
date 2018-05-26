VERSION 5.00
Begin VB.Form frmLogicGates 
   BackColor       =   &H00800080&
   Caption         =   "Logic Gates"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboLogicGates 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton cmdReturn3 
      Caption         =   "&Return"
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
      Left            =   5640
      TabIndex        =   0
      Top             =   4680
      Width           =   1575
   End
End
Attribute VB_Name = "frmLogicGates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboLogicGates_Click()
    cboLogicGates.AddItem "AND GATE"
End Sub

Private Sub cmdReturn3_Click()
    Unload frmLogicGates
End Sub

Private Sub Combo1_Change()

End Sub
