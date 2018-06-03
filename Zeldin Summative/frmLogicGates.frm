VERSION 5.00
Begin VB.Form frmLogicGates 
   BackColor       =   &H00800080&
   Caption         =   "Logic Gates"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraInput2 
      Caption         =   "Input 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   3960
      Width           =   1605
      Begin VB.OptionButton optTRUE 
         Caption         =   "TRUE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optFALSE 
         Caption         =   "FALSE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame fraInput 
      Caption         =   "Input 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1605
      Begin VB.OptionButton optFALSE 
         Caption         =   "FALSE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optTRUE 
         Caption         =   "TRUE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.ComboBox cboLogicGates 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3720
      Width           =   2160
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "&Return to Main Form"
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
      Left            =   6480
      TabIndex        =   0
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Image imgGate1 
      Height          =   1695
      Index           =   5
      Left            =   4200
      Picture         =   "frmLogicGates.frx":0000
      Stretch         =   -1  'True
      Top             =   2025
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate1 
      Height          =   1695
      Index           =   4
      Left            =   4200
      Picture         =   "frmLogicGates.frx":2F22
      Stretch         =   -1  'True
      Top             =   2025
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate1 
      Height          =   1695
      Index           =   3
      Left            =   4200
      Picture         =   "frmLogicGates.frx":5E44
      Stretch         =   -1  'True
      Top             =   2025
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate1 
      Height          =   1695
      Index           =   2
      Left            =   4200
      Picture         =   "frmLogicGates.frx":8D66
      Stretch         =   -1  'True
      Top             =   2025
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate1 
      Height          =   1695
      Index           =   1
      Left            =   4200
      Picture         =   "frmLogicGates.frx":BC88
      Stretch         =   -1  'True
      Top             =   2025
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate1 
      Height          =   1695
      Index           =   0
      Left            =   4200
      Picture         =   "frmLogicGates.frx":EBAA
      Stretch         =   -1  'True
      Top             =   2025
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label lblTrueorFalse1 
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6400
      TabIndex        =   2
      Top             =   2625
      Width           =   1335
   End
End
Attribute VB_Name = "frmLogicGates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Zeldin
Option Explicit
Dim Gate(0 To 5) As String
Dim GateOutput(0 To 2) As Boolean
Private Function calculateLogic() As Boolean
    Select Case cboLogicGates
        Case "AND"
            calculateLogic = optTRUE(0) And optTRUE(1)
        Case "OR"
            calculateLogic = optTRUE(0) Or optTRUE(1)
        Case "XOR"
            calculateLogic = optTRUE(0) Xor optTRUE(1)
        Case "NAND"
            calculateLogic = Not (optTRUE(0) And optTRUE(1))
        Case "NOR"
            calculateLogic = Not (optTRUE(0) Or optTRUE(1))
        Case "XNOR"
            calculateLogic = Not (optTRUE(0) Xor optTRUE(1))
    End Select
End Function
Private Sub inputChanged(Index As Integer)
    Dim i
    If cboLogicGates = "" Then
        ' No gate is selected - nothing to do
        lblTrueorFalse1.Visible = False
        Exit Sub
    End If
        
    For i = 0 To 1 'UBound(optTRUE)
        If Not (optTRUE(i) Or optFALSE(i)) Then
            lblTrueorFalse1.Visible = False
            Exit Sub
        End If
    Next i
    lblTrueorFalse1 = calculateLogic
    lblTrueorFalse1.Visible = True
End Sub
Private Sub cboLogicGates_Click()
    Dim i As Integer
    For i = 0 To UBound(Gate)
        imgGate1(i).Visible = cboLogicGates = Gate(i)
    Next i
    ' Update lblTrueorFalse1
    inputChanged (0)
End Sub

Private Sub cmdReturn_Click()
    Unload frmLogicGates
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Gate(0) = "AND"
    Gate(1) = "OR"
    Gate(2) = "XOR"
    Gate(3) = "NAND"
    Gate(4) = "NOR"
    Gate(5) = "XNOR"
    
    For i = 0 To UBound(Gate)
        cboLogicGates.AddItem (Gate(i))
        imgGate1(i).Visible = False
    Next i
End Sub

Private Sub optTRUE_Click(Index As Integer)
    inputChanged (Index)
End Sub

Private Sub optFALSE_Click(Index As Integer)
    inputChanged (Index)
End Sub

