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
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   1560
         X2              =   1560
         Y1              =   1440
         Y2              =   960
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
      Left            =   7800
      TabIndex        =   0
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Line Line13 
      X1              =   9000
      X2              =   10440
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line12 
      X1              =   9000
      X2              =   9000
      Y1              =   2880
      Y2              =   6480
   End
   Begin VB.Line Line11 
      X1              =   7920
      X2              =   9000
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line10 
      X1              =   7680
      X2              =   9000
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line9 
      X1              =   3240
      X2              =   4080
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line8 
      X1              =   3240
      X2              =   3240
      Y1              =   6240
      Y2              =   7800
   End
   Begin VB.Line Line7 
      X1              =   1920
      X2              =   3240
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line6 
      X1              =   2040
      X2              =   3240
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line5 
      X1              =   3240
      X2              =   4200
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line4 
      X1              =   3240
      X2              =   3240
      Y1              =   2160
      Y2              =   3840
   End
   Begin VB.Line Line3 
      X1              =   1920
      X2              =   3240
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line2 
      X1              =   1800
      X2              =   3240
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblLogicGatesTitle 
      BackColor       =   &H00800080&
      Caption         =   "Welcome to the Logic Gates Simulator! Please chose a gate and an input!"
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
      Left            =   4200
      TabIndex        =   19
      Top             =   240
      Width           =   7095
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
      Left            =   6480
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

