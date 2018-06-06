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
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00000000&
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
      BackColor       =   &H00C0C0C0&
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
      Left            =   960
      TabIndex        =   0
      Top             =   6000
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   1215
      Left            =   600
      TabIndex        =   7
      Top             =   4680
      Width           =   3015
      WordWrap        =   -1  'True
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
Dim numberValid As Boolean
Dim base As Integer
Dim baseValid As Boolean
Dim convertedNumber As String
Dim stepValue 'the value of the current step
Private Function setDecimalNum(str As String) As Boolean
    'The purpose of the function is to check if the input is between 0 and 255
    Dim n As Integer
    n = Val(str)
    If IsNumeric(str) And n >= 0 And n < 256 Then
        lblResult.Caption = ""
        originalNumber = n
        stepValue = n
        numberValid = True
        setDecimalNum = True
    Else
        lblResult.Caption = "Input decimal number between 0 and 255."
        numberValid = False
        setDecimalNum = False
    End If
End Function
Private Function setBase(str As String) As Boolean
    'Makes sure that base is valid
    Dim n As Integer
    n = Val(str)
    If IsNumeric(str) And n >= 2 And n <= 9 Or n = 16 Then
        base = n
        baseValid = True
        lblResult.Caption = ""
        setBase = True
    Else
        lblResult.Caption = "Enter base as a decimal number between 2 and 9 or 16"
        baseValid = False
        setBase = False
    End If
End Function
Private Sub setCmdButtons(enabledOrNot As Boolean)
    'Enables and disables controls
    cmdShowAnswer.Enabled = enabledOrNot
    cmdStepThrough.Enabled = enabledOrNot
End Sub

Private Sub prepCmd()
    'Checks if all inputs have been provided and enables command buttons
    lblSteps.Caption = ""
    lblResult.Caption = ""
    If baseValid And numberValid Then
        stepValue = originalNumber
        setCmdButtons (True)
    Else
        setCmdButtons (False)
    End If
End Sub
Private Function digit2str(d As Integer) As String
    'This part converts number to string that represents hex characters
    Select Case d
        Case 0 To 9
            digit2str = str(d)
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
Private Function stepThroughConvertor() As String
    'This peforms one step of number conversion. This function is called from Show Answer and Step Through buttons.
    Dim digit As Integer
    digit = stepValue Mod base
    stepValue = Int(stepValue / base)
    stepThroughConvertor = digit2str(digit)
End Function
Private Sub numberConvertor()
    'Converts numbers
    Dim digit As Integer
    Dim oldStepValue As Integer 'for coordination with stepThrough
    If originalNumber = 0 Then
        lblResult.Caption = "0"
        Exit Sub
    End If
    
    lblResult.Caption = ""
    oldStepValue = stepValue
    stepValue = originalNumber
    
    While stepValue > 0 'It calls the stepthrough convertor function until the number has been fully converted
        lblResult.Caption = stepThroughConvertor() & lblResult.Caption
    Wend
    
    'reset stepValue back so that stepThrough works
    stepValue = oldStepValue
End Sub
Private Sub initForm()
    'Makes the form blank when initally loaded or when clear button is clicked
    txtBase.Text = ""
    txtDecimalNumber.Text = ""
    lblSteps.Caption = ""
    lblResult.Caption = ""
    vsbBase.Value = 0
    baseValid = False
    numberValid = False
    originalNumber = 0
    base = 0
    prepCmd
End Sub
Private Sub cmdClear_Click()
    'Clears the form
    initForm
End Sub
Private Sub cmdReturn_Click()
    'returns back to main
    Unload frmRemainderMethod
End Sub

Private Sub cmdShowAnswer_Click()
    'Allows only one click
    cmdShowAnswer.Enabled = False
    numberConvertor
End Sub

Private Sub cmdStepThrough_Click()
    Dim digit As String
    Dim prevStep As Integer
    prevStep = stepValue
    digit = stepThroughConvertor() 'stepValue is changed now
    lblSteps.Caption = lblSteps.Caption & str(prevStep) & "/" & str(base) & "=" & str(stepValue) & " R" & digit & vbCrLf
    If stepValue = 0 Then
        'We are done, no more steps should be shown
        cmdStepThrough.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    initForm
End Sub

Private Sub txtBase_Change()
    setCmdButtons (False)
    If setBase(txtBase.Text) Then
        prepCmd
    End If
End Sub

Private Sub txtDecimalNumber_Change()
    setCmdButtons (False)
    If setDecimalNum(txtDecimalNumber) Then
        prepCmd
    End If
End Sub
Private Sub vsbBase_Change()
    setCmdButtons (False)
    txtBase.Text = str(vsbBase.Value)
    If setBase(txtBase.Text) Then
        prepCmd
    End If
End Sub
