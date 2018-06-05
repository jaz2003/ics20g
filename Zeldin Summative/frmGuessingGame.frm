VERSION 5.00
Begin VB.Form frmGuessingGame 
   BackColor       =   &H000000FF&
   Caption         =   "Guessing Game"
   ClientHeight    =   9315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9720
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraGuessLimit 
      BackColor       =   &H000000FF&
      Caption         =   "Guess Limit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   6600
      TabIndex        =   21
      Top             =   3600
      Width           =   2295
      Begin VB.OptionButton opt5Guesses 
         BackColor       =   &H000000FF&
         Caption         =   "5 Guesses"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton opt8Guesses 
         BackColor       =   &H000000FF&
         Caption         =   "8 Guesses"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton opt12Guesses 
         BackColor       =   &H000000FF&
         Caption         =   "12 Guesses"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   1815
      End
   End
   Begin VB.Timer tmrTimeLimit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   480
   End
   Begin VB.Frame fraTimeLimit 
      BackColor       =   &H000000FF&
      Caption         =   "Time Limit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   6600
      TabIndex        =   13
      Top             =   600
      Width           =   2295
      Begin VB.OptionButton opt25Seconds 
         BackColor       =   &H000000FF&
         Caption         =   "25 Seconds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   1815
      End
      Begin VB.OptionButton opt15Seconds 
         BackColor       =   &H000000FF&
         Caption         =   "15 Seconds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton opt7Seconds 
         BackColor       =   &H000000FF&
         Caption         =   "7 Seconds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.TextBox txtMax 
      BackColor       =   &H000000FF&
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtMin 
      BackColor       =   &H000000FF&
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSubmitGuess 
      Caption         =   "&Guess"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox txtGuess 
      BackColor       =   &H000000FF&
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
      Left            =   2640
      TabIndex        =   4
      Top             =   4680
      Width           =   1695
   End
   Begin VB.HScrollBar hsbGuess 
      Height          =   255
      Left            =   1200
      Max             =   100
      TabIndex        =   3
      Top             =   5400
      Width           =   4335
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Press to Start!"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn1 
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
      Height          =   735
      Left            =   4800
      TabIndex        =   0
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label lblguessCountTitle 
      BackColor       =   &H000000FF&
      Caption         =   "Guess Count:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblMax 
      Caption         =   "Max:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   20
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblMin 
      Caption         =   "Min:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   19
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblTitletimeLeft 
      BackColor       =   &H000000FF&
      Caption         =   "Time Left:"
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
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lbltimeLeft 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblIntro 
      BackColor       =   &H000000FF&
      Caption         =   "Welcome to the Guesing Game: Enter a max and a min, a time limit, maximum amount of guesses.  Have fun ! :) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1440
      TabIndex        =   12
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label lblguessCount 
      BackColor       =   &H000000FF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblTooHigh 
      BackStyle       =   0  'Transparent
      Caption         =   "Too High"
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
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblTooLow 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Too Low"
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
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblResult 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1920
      TabIndex        =   5
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label lblTitle2 
      Caption         =   "The Number Guessing Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmGuessingGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Zeldin
Option Explicit
Dim num As Integer
Dim guess As Integer
Dim guessCount As Integer
Dim max As Integer
Dim min As Integer
Dim timeSelected As Boolean
Dim rangeSelected As Boolean
Dim guessLimitSelected As Boolean
Dim guessLimit As Integer
Dim timeLeft As Integer
Private Sub initControls()
    hsbGuess.Enabled = False
    txtGuess.Enabled = False
    lblTooHigh.Enabled = False
    lblTooLow.Enabled = False
    lblTooHigh.ForeColor = RGB(200, 0, 0)
    lblTooLow.ForeColor = RGB(200, 0, 0)
    lblguessCount.Enabled = False
    cmdSubmitGuess.Enabled = False
    tmrTimeLimit.Enabled = False
    lbltimeLeft.Enabled = False
    lbltimeLeft.Enabled = False
    lblTitletimeLeft.Enabled = False
    opt5Guesses.Value = False
    opt8Guesses.Value = False
    opt12Guesses.Value = False
    opt7Seconds.Value = False
    opt15Seconds.Value = False
    opt25Seconds.Value = False
    guessCount = 0
End Sub
Private Sub cmdReturn1_Click()
    Unload frmGuessingGame
End Sub
Private Sub cmdStart_Click()
    max = Val(txtMax.Text)
    min = Val(txtMin.Text)
      
    'Checks to make sure that max is greater than min
    If max < min Then
        lblResult.Caption = "Error, your max is less than your min"
        rangeSelected = False
    Else
        rangeSelected = True
    End If
     
    'Checks to make sure that user selected guess limit
    If opt5Guesses.Value = False And opt8Guesses.Value = False And opt12Guesses.Value = False Then
        guessLimitSelected = False
        lblResult.Caption = "Error! You have not selected a guess limit!"
    ElseIf opt5Guesses.Value = True Then
        guessLimitSelected = True
        guessLimit = 5
    ElseIf opt8Guesses.Value = True Then
        guessLimitSelected = True
        guessLimit = 8
    ElseIf opt12Guesses.Value = True Then
        guessLimitSelected = True
        guessLimit = 12
    End If
    
    'Makes sure that user selects time setting
    If opt7Seconds.Value = False And opt15Seconds.Value = False And opt25Seconds.Value = False Then
        lblResult.Enabled = True
        lblResult.Caption = "You have not selected a time limit."
        timeSelected = False
    ElseIf opt7Seconds.Value = True Then
        timeLeft = 7
        timeSelected = True
    ElseIf opt15Seconds.Value = True Then
        timeLeft = 15
        timeSelected = True
    Else
        timeLeft = 25
        timeSelected = True
    End If
 
           
    'Creates conditions
    If rangeSelected = True And timeSelected = True And guessLimitSelected = True Then
        num = Int(Rnd * max + min)
        hsbGuess.min = min
        hsbGuess.max = max
        hsbGuess.Enabled = True
        txtGuess.Enabled = True
        txtGuess = ""
        guessCount = 0
        num = Int(Rnd * (max - min) + min) 'Generates random number
        lblResult.Caption = ""
        lblTooHigh.Enabled = True
        lblTooLow.Enabled = True
        tmrTimeLimit.Enabled = True
        lbltimeLeft.Enabled = True
        lbltimeLeft.Caption = str(timeLeft)
        lblguessCount.Enabled = False
        cmdSubmitGuess.Enabled = True
        lblTitletimeLeft.Enabled = True
        fraTimeLimit.Caption = "Time Limit"
        fraGuessLimit.Caption = "Guess Limit"
        txtGuess.Enabled = True
        lblResult.Enabled = True
        lblResult.Caption = ""
        lblguessCount.Caption = Val(0)
        guess = min - 1 'So it's not a valid guess
    End If
End Sub
Private Sub cmdSubmitGuess_Click()
    If guess = num And timeLeft > 0 And guessLimit > guessCount Then
        initControls
        cmdStart.Caption = "Restart"
        lblResult.Caption = "Correct! You have won :)"
        lblTooHigh.ForeColor = RGB(200, 0, 0)
        lblTooLow.ForeColor = RGB(200, 0, 0)
    ElseIf guess < num And timeLeft > 0 And guessLimit > guessCount Then
        lblTooHigh.ForeColor = RGB(200, 0, 0)
        lblTooLow.ForeColor = RGB(0, 0, 0)
    ElseIf guess > num And timeLeft > 0 And guessLimit > guessCount Then
       lblTooLow.ForeColor = RGB(200, 0, 0)
       lblTooHigh.ForeColor = RGB(0, 0, 0)
    End If
    
    'Updates display
    guessCount = guessCount + 1
    lblguessCount.Caption = guessCount
End Sub

Private Sub Form_Load()
    initControls
End Sub

Private Sub hsbGuess_Change()
    'Determines what user guess is
    guess = hsbGuess.Value
    
    'Changes text in the text box
    txtGuess.Text = str(guess)
End Sub

Private Sub lblguessCount_Change()
    'What happens if user rans out of guesses
    If guessCount = guessLimit Then
        tmrTimeLimit.Enabled = False
        cmdStart.Caption = "Restart"
        initControls
        lblResult.Caption = "You lost. You ran out of guesses :(."
    End If
End Sub
Private Sub tmrTimeLimit_Timer()
    timeLeft = timeLeft - 1
    lbltimeLeft.Caption = str(timeLeft)
    
    If timeLeft = 0 Then
        tmrTimeLimit.Enabled = False
    End If
    
    If guess <> num And timeLeft = 0 Then
        initControls
        lblResult.Caption = "You ran out of time :(. This means that you lost."
        cmdStart.Caption = "Restart"
    ElseIf guess = num Then
        tmrTimeLimit.Enabled = False
    End If
End Sub

Private Sub txtGuess_Validate(cancel As Boolean)
    cancel = True 'Assume something is wrong
    If IsNumeric(txtGuess.Text) Then
        'Determines what user guess is
        guess = Val(txtGuess.Text)
        If guess < min Then
            lblResult.Caption = "Value must be >= " & str(min)
        ElseIf guess > max Then
            lblResult.Caption = "Value must be <= " & str(max)
        Else
            cancel = False
            lblResult.Caption = "" ' Erase whatever was displayed there before
            'Updates scrollbar
            hsbGuess.Value = guess
        End If
    Else
        lblResult.Caption = "Value must be numeric."
    End If
  End Sub
