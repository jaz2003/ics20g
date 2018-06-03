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
   Begin VB.Frame fraGate1B 
      Caption         =   "Gate 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   1695
      Begin VB.OptionButton optTRUE2 
         Caption         =   "TRUE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optFALSE2 
         Caption         =   "FALSE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Frame fraGateA 
      Caption         =   "Gate 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
      Begin VB.OptionButton optFALSE1 
         Caption         =   "FALSE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton optTRUE1 
         Caption         =   "TRUE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.ComboBox cboLogicGates 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3720
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
      Left            =   6480
      TabIndex        =   0
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Image imgNandGate1 
      Height          =   1695
      Index           =   1
      Left            =   0
      Picture         =   "frmLogicGates.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgNandGate1 
      Height          =   1695
      Index           =   0
      Left            =   4320
      Picture         =   "frmLogicGates.frx":2F22
      Stretch         =   -1  'True
      Top             =   1920
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgXorGate1 
      Height          =   1695
      Left            =   4320
      Picture         =   "frmLogicGates.frx":5E44
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgXnorGate1 
      Height          =   1695
      Left            =   4320
      Picture         =   "frmLogicGates.frx":8D66
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgOrGate1 
      Height          =   1695
      Left            =   4320
      Picture         =   "frmLogicGates.frx":BC88
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgNorGate1 
      Height          =   1695
      Left            =   4440
      Picture         =   "frmLogicGates.frx":EBAA
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgAndGate1 
      Height          =   1695
      Left            =   4440
      Picture         =   "frmLogicGates.frx":11ACC
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label lblTrueorFalse1 
      BackColor       =   &H00800080&
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   2520
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
Dim GateName(0 To 5) As String
Dim GateImage(0 To 5) As Image
Dim GateOutput(0 To 2) As Boolean
Private Sub cboLogicGates_Click()
    Dim i As Integer
    For i = 0 To UBound(GateName)
        If (GateName(i) = cboLogicGates) Then
            GateVisible(i) = True
        Else
            GateVisible(i) = True
        End If
        'GateVisible(i) = (GateName(i) = cboLogicGates)
    Next i
    
    If cboLogicGates = "AND GATE" Then
       imgAndGate1.Visible = True
       imgNorGate1.Visible = False
       imgOrGate1.Visible = False
       imgXorGate1.Visible = False
       imgXnorGate1.Visible = False
       imgNandGate1.Visible = False
    ElseIf cboLogicGates = "OR GATE" Then
       imgAndGate1.Visible = False
       imgNorGate1.Visible = False
       imgOrGate1.Visible = False
       imgXorGate1.Visible = False
       imgXnorGate1.Visible = True
       imgNandGate1.Visible = False
    ElseIf cboLogicGates = "XOR GATE" Then
       imgAndGate1.Visible = False
       imgNorGate1.Visible = False
       imgOrGate1.Visible = False
       imgXorGate1.Visible = True
       imgXnorGate1.Visible = False
       imgNandGate1.Visible = False
    ElseIf cboLogicGates = "NAND GATE" Then
       Call NandGate1
       imgAndGate1.Visible = False
       imgNorGate1.Visible = False
       imgOrGate1.Visible = False
       imgXorGate1.Visible = False
       imgXnorGate1.Visible = False
       imgNandGate1.Visible = True
    ElseIf cboLogicGates = "XNOR GATE" Then
       imgAndGate1.Visible = False
       imgNorGate1.Visible = False
       imgOrGate1.Visible = False
       imgXorGate1.Visible = False
       imgXnorGate1.Visible = True
       imgNandGate1.Visible = False
    ElseIf cboLogicGates = "NOR GATE" Then
       imgAndGate1.Visible = False
       imgNorGate1.Visible = True
       imgOrGate1.Visible = False
       imgXorGate1.Visible = False
       imgXnorGate1.Visible = False
       imgNandGate1.Visible = False
    End If
    
End Sub

Private Sub cmdReturn3_Click()
    Unload frmLogicGates
End Sub

Private Sub Form_Load()
    Dim i As Integer
    GateName(0) = "AND"
    GateImage(0) = imgAndGate1.Picture
    GateName(1) = "OR"
    GateImage(1) = imgOrGate1
    GateName(2) = "XOR"
    GateImage(2) = imgXorGate1
    GateName(3) = "NAND"
    GateImage(3) = imgNandGate1
    GateName(4) = "NOR"
    GateImage(4) = imgNorGate1
    GateName(5) = "XNOR"
    GateImage(5) = imgXnorGate1
    
    For i = 0 To UBound(GateName)
        cboLogicGates.AddItem (GateName(i))
        GateImage(i) = False
    Next i
        
    cboLogicGates.AddItem "AND GATE"
    cboLogicGates.AddItem "OR GATE"
    cboLogicGates.AddItem "XOR GATE"
    cboLogicGates.AddItem "NAND GATE"
    cboLogicGates.AddItem "NOR GATE"
    cboLogicGates.AddItem "XNOR GATE"
End Sub
Private Sub NandGate1()
    If optTrue1.Value = True And optTrue2.Value = True Then
        lblTrueorFalse1.Caption = "FALSE"
    Else: lblTrueorFalse1.Caption = "TRUE"
    End If
End Sub

