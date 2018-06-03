VERSION 5.00
Begin VB.Form frmLogicGates 
   BackColor       =   &H00800080&
   Caption         =   "Logic Gates"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   15060
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboLogicGates 
      Height          =   315
      Index           =   2
      Left            =   10440
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   5520
      Width           =   2160
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
      Index           =   1
      Left            =   360
      TabIndex        =   13
      Top             =   5280
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
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   480
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
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
   End
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
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   6960
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
         Index           =   2
         Left            =   240
         TabIndex        =   12
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
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.ComboBox cboLogicGates 
      Height          =   315
      Index           =   1
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   7440
      Width           =   2160
   End
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
      Top             =   3120
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
      Index           =   0
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
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblTrueorFalse 
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
      Index           =   2
      Left            =   12645
      TabIndex        =   18
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   17
      Left            =   10440
      Picture         =   "frmLogicGates.frx":0000
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   16
      Left            =   10440
      Picture         =   "frmLogicGates.frx":2F22
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   15
      Left            =   10440
      Picture         =   "frmLogicGates.frx":5E44
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   14
      Left            =   10440
      Picture         =   "frmLogicGates.frx":8D66
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   13
      Left            =   10440
      Picture         =   "frmLogicGates.frx":BC88
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   12
      Left            =   10440
      Picture         =   "frmLogicGates.frx":EBAA
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label lblTrueorFalse 
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
      Index           =   1
      Left            =   6240
      TabIndex        =   16
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   11
      Left            =   4080
      Picture         =   "frmLogicGates.frx":11ACC
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   10
      Left            =   4080
      Picture         =   "frmLogicGates.frx":149EE
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   9
      Left            =   4080
      Picture         =   "frmLogicGates.frx":17910
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   8
      Left            =   4080
      Picture         =   "frmLogicGates.frx":1A832
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   7
      Left            =   4080
      Picture         =   "frmLogicGates.frx":1D754
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   6
      Left            =   4080
      Picture         =   "frmLogicGates.frx":20676
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   5
      Left            =   4200
      Picture         =   "frmLogicGates.frx":23598
      Stretch         =   -1  'True
      Top             =   2025
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   4
      Left            =   4200
      Picture         =   "frmLogicGates.frx":264BA
      Stretch         =   -1  'True
      Top             =   2025
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   3
      Left            =   4200
      Picture         =   "frmLogicGates.frx":293DC
      Stretch         =   -1  'True
      Top             =   2025
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   2
      Left            =   4200
      Picture         =   "frmLogicGates.frx":2C2FE
      Stretch         =   -1  'True
      Top             =   2025
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   1
      Left            =   4200
      Picture         =   "frmLogicGates.frx":2F220
      Stretch         =   -1  'True
      Top             =   2025
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgGate 
      Height          =   1695
      Index           =   0
      Left            =   4200
      Picture         =   "frmLogicGates.frx":32142
      Stretch         =   -1  'True
      Top             =   2025
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label lblTrueorFalse 
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
      Index           =   0
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
Private Const numGateOps As Integer = 6
Private Const numGates As Integer = 3
Private Const numInputs = 2 * numGates
Dim Gate(0 To numGateOps - 1) As String
Dim inputVal(0 To numInputs - 1) As Boolean
Dim inputDefined(0 To numInputs - 1) As Boolean

Private Function calculateLogic(gateOp As String, input1 As Boolean, input2 As Boolean) As Boolean
        'Calculates gate output for any gate (true or false)
    Select Case gateOp
        Case "AND"
            calculateLogic = input1 And input2
        Case "OR"
            calculateLogic = input1 Or input2
        Case "XOR"
            calculateLogic = input1 Xor input2
        Case "NAND"
            calculateLogic = Not (input1 And input2)
        Case "NOR"
            calculateLogic = Not (input1 Or input2)
        Case "XNOR"
            calculateLogic = Not (input1 Xor input2)
    End Select
End Function

Private Sub updateGate(gateIdx As Integer)
    Dim nextInputIdx As Integer 'where output of this gate is connected to
    Dim gateOutput As Boolean
    
    'Makes sure that gate output is visible only if both gate inputs have been defined
    'and the gate type is selected
    If Not (inputDefined(2 * gateIdx) And inputDefined(2 * gateIdx + 1) And cboLogicGates(gateIdx) <> "") Then
        lblTrueorFalse(gateIdx).Visible = False
        Exit Sub
    End If
    gateOutput = calculateLogic(cboLogicGates(gateIdx), inputVal(2 * gateIdx), inputVal(2 * gateIdx + 1))
    lblTrueorFalse(gateIdx) = gateOutput
    lblTrueorFalse(gateIdx).Visible = True
    
    Select Case gateIdx
        Case 0
            nextInputIdx = 4
        Case 1
            nextInputIdx = 5
        Case Else
            Exit Sub
    End Select
    
    inputDefined(nextInputIdx) = True
    inputVal(nextInputIdx) = gateOutput
    updateGate (2)
 
End Sub
Private Sub inputChanged(idx As Integer)
    Dim gateIdx As Integer
    gateIdx = Int(idx / 2) ' gate(0) has input(0) and input(1), gate(1) has input(2) and input(3)
    inputDefined(idx) = optTRUE(idx) Or optFALSE(idx)
    inputVal(idx) = optTRUE(idx)
    updateGate (gateIdx)
End Sub
Private Sub cboLogicGates_Click(gateIdx As Integer)
    Dim i As Integer
    For i = 0 To UBound(Gate)
        imgGate((UBound(Gate) + 1) * gateIdx + i).Visible = cboLogicGates(gateIdx) = Gate(i)
    Next i
    ' Update lblTrueorFalse if both inputs are defined
    updateGate (gateIdx)
End Sub

Private Sub cmdReturn_Click()
    Unload frmLogicGates
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim j As Integer
    
    Gate(0) = "AND"
    Gate(1) = "OR"
    Gate(2) = "XOR"
    Gate(3) = "NAND"
    Gate(4) = "NOR"
    Gate(5) = "XNOR"
    
    For j = 0 To numGates - 1 'UBound(cboLogicGates)
        For i = 0 To UBound(Gate)
            cboLogicGates(j).AddItem (Gate(i))
        Next i
    Next j
    
    'Initalizes the input defined array
    For i = 0 To UBound(inputDefined)
        inputDefined(i) = False
    Next i
End Sub

Private Sub optTRUE_Click(Index As Integer)
    inputChanged (Index)
End Sub

Private Sub optFALSE_Click(Index As Integer)
    inputChanged (Index)
End Sub

