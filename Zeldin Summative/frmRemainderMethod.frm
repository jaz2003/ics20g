VERSION 5.00
Begin VB.Form frmRemainderMethod 
   BackColor       =   &H00404040&
   Caption         =   "Remainder Method"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
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
      Left            =   4920
      TabIndex        =   11
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton cmdStepThrough 
      Caption         =   "Step Through"
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
      Left            =   4920
      TabIndex        =   9
      Top             =   5520
      Width           =   2775
   End
   Begin VB.CommandButton cmdShowAnswer 
      Caption         =   "Show Answer"
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
      Left            =   960
      TabIndex        =   6
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox txtDecimalNumber 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2880
      Width           =   735
   End
   Begin VB.VScrollBar vsbBase 
      Height          =   615
      Left            =   2640
      Max             =   0
      Min             =   16
      TabIndex        =   2
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtBase 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
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
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00404040&
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
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label lblRemainderMethodTitle 
      BackColor       =   &H00404040&
      Caption         =   "Remainder Method"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   3720
      TabIndex        =   10
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label lblSteps 
      BackColor       =   &H00808080&
      Height          =   3975
      Left            =   4080
      TabIndex        =   8
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label lblResult 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
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
      Left            =   600
      TabIndex        =   7
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Label lblDecimalNumber 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Decimal Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   480
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblBase 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Base"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "frmRemainderMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Zeldin
'Option Explicit
Dim originalNumber As Integer 'the number that the user wrote initally
Dim base As Integer
Dim convertedNumber As String
Dim stepValue 'the value of the current step
Private Function digit2str(d As Integer)
    Select Case d
        Case 0 To 9
            digit2str = Str(d)
        Case 10
            digit2str = "A"
        Case 11
            digit2str = "B"
        Case 12
            digit2str = "C"
        Case 13
            digit2str = "D"
        Case 14
            digit2str = "E"
        Case 15
            digit2str = "F"
    End Select
End Function


Private Sub cmdClear_Click()
    'Enables controls
    cmdShowAnswer.Enabled = True
    cmdStepThrough.Enabled = True
    txtBase.Text = ""
    txtDecimalNumber.Text = ""
    lblSteps.Caption = ""
    lblResult.Caption = ""
    vsbBase.Value = 0
    originalNumber = 0
    stepValue = originalNumber
End Sub
Private Sub cmdReturn_Click()
    Unload frmRemainderMethod
End Sub
Private Sub numberConvertor()
    'Converts numbers
    Dim digit As Integer
    If originalNumber = 0 Then
        lblResult.Caption = "0"
        Exit Sub
    End If
    
    stepValue = originalNumber
    lblResult.Caption = ""
    While stepValue > 0
        digit = stepValue Mod base
        stepValue = Int(stepValue / base)
        lblResult.Caption = digit2str(digit) & lblResult.Caption
    Wend
End Sub

Private Sub cmdShowAnswer_Click()
    'Allows only one click
    cmdShowAnswer.Enabled = False
    
    'Determines that input is valid
        If base < 2 Or (base > 9 And base <> 16) Then
            lblResult.Caption = "You have selected an invalid base"
        Else: Call numberConvertor
        End If
        End Sub

Private Sub txtBase_Change()
    base = Val(txtBase.Text)
End Sub

Private Sub txtDecimalNumber_Validate(keepFocus As Boolean)
    originalNumber = Val(txtDecimalNumber)
    If IsNumeric(txtDecimalNumber) And originalNumber >= 0 And originalNumber < 256 Then
        keepFocus = False
    Else
        keepFocus = True
        lblResult.Caption = "Input decimal number between 0 and 255."
    End If
End Sub

Private Sub vsbBase_Change()
    txtBase.Text = Str(vsbBase.Value)
    base = vsbBase.Value
End Sub
