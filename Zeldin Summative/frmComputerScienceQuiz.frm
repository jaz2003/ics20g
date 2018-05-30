VERSION 5.00
Begin VB.Form frmComputerScienceQuiz 
   BackColor       =   &H00800000&
   Caption         =   "Computer Science Quiz"
   ClientHeight    =   10170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   ScaleHeight     =   10170
   ScaleWidth      =   12510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNewQuestions 
      Caption         =   "New Questions"
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
      Left            =   720
      TabIndex        =   19
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   5040
      TabIndex        =   12
      Top             =   5640
      Width           =   1935
      Begin VB.OptionButton optTrue3 
         Caption         =   "True:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optFalse3 
         Caption         =   "False:"
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
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   5040
      TabIndex        =   6
      Top             =   3480
      Width           =   1935
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1935
         Begin VB.OptionButton optTrue2 
            Caption         =   "True:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optFalse2 
            Caption         =   "False:"
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
            TabIndex        =   10
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "True:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "False:"
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
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame fraQuestion1 
      Height          =   1455
      Left            =   5040
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
      Begin VB.OptionButton optFalse1 
         Caption         =   "False:"
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
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optTrue1 
         Caption         =   "True:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn4 
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
      Height          =   855
      Left            =   5520
      TabIndex        =   0
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Image imgWrong3 
      Height          =   1905
      Left            =   7320
      Picture         =   "frmComputerScienceQuiz.frx":0000
      Top             =   5760
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image imgWrong2 
      Height          =   1905
      Left            =   7320
      Picture         =   "frmComputerScienceQuiz.frx":099C
      Top             =   3480
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image imgWrong1 
      Height          =   1905
      Left            =   7320
      Picture         =   "frmComputerScienceQuiz.frx":1338
      Top             =   1440
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image imgRight3 
      Height          =   1590
      Left            =   7320
      Picture         =   "frmComputerScienceQuiz.frx":1CD4
      Top             =   5760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image imgRight2 
      Height          =   1590
      Left            =   7320
      Picture         =   "frmComputerScienceQuiz.frx":2469
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image imgRight1 
      Height          =   1590
      Left            =   7320
      Picture         =   "frmComputerScienceQuiz.frx":2BFE
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblComputerQuizResult 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      TabIndex        =   18
      Top             =   8760
      Width           =   6975
   End
   Begin VB.Label lblQuestion2 
      Height          =   1095
      Left            =   480
      TabIndex        =   17
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label lblQuestion3 
      Height          =   1095
      Left            =   480
      TabIndex        =   16
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label lblQuestion1 
      Height          =   1095
      Left            =   480
      TabIndex        =   15
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lblbllComputerQuizTitle 
      Caption         =   "THE COMPUTER QUIZ:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmComputerScienceQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Zeldin
Option Explicit
Dim question(1 To 20) As String
Dim answer(1 To 20) As Boolean
Dim randomQuestion(1 To 20) As Integer
' Index of first of the three current questions in randomQuestions array
Dim currentQuestionStart As Integer

' Generate array of numbers from 1 to NumItems in random order
' Inspired by http://www.vb-helper.com/howto_randomize_array.html
Private Sub randomizeQuestions()
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim tmp As Integer
    n = UBound(randomQuestion)
    ' initialize randomQuestion array
    For i = 1 To n
        randomQuestion(i) = i
    Next i
    
    ' shuffle randomQuestion
    For i = 1 To n - 1
        ' Pick a random entry.
        j = Int((n - i + 1) * Rnd + i)
        ' Swap the numbers.
        tmp = randomQuestion(i)
        randomQuestion(i) = randomQuestion(j)
        randomQuestion(j) = tmp
    Next i
End Sub
Private Sub generateQuestions()
    Dim n As Integer
    n = UBound(randomQuestion)
    currentQuestionStart = currentQuestionStart + 3
    If currentQuestionStart > n Then
        randomizeQuestions
        currentQuestionStart = 1
    End If
    
    lblQuestion1.Caption = question(randomQuestion(currentQuestionStart))
    lblQuestion2.Caption = question(randomQuestion(1 + currentQuestionStart Mod n))
    lblQuestion3.Caption = question(randomQuestion(1 + (currentQuestionStart + 1) Mod n))
End Sub

Private Sub cmdNewQuestions_Click()
    generateQuestions
End Sub

Private Sub cmdReturn4_Click()
    Unload frmComputerScienceQuiz
End Sub

Private Sub cmdSubmit_Click()
    Dim questionAnswered(3) As Boolean
    
    
    questionAnswered(1) = optTrue1 Or optFalse1
    questionAnswered(2) = optTrue2 Or optFalse2
    questionAnswered(3) = optTrue3 Or optFalse3
    
    'Determines if user answer is correct for question 1
    If optTrue1.Value = True Then
        imgWrong1.Visible = True
        imgRight1.Visible = False
    ElseIf optFalse1.Value = True Then
        imgRight1.Visible = True
        imgWrong1.Visible = False
    End If
    
    'Determines if user answer is correct for question 2
    If optTrue2.Value = True Then
        imgRight2.Visible = True
        imgWrong2.Visible = False
    ElseIf optFalse2.Value = True Then
        imgWrong2.Visible = True
        imgRight2.Visible = False
    End If
    
    'Determines that us
    
     'Determines if user answer is correct for question 3
     If optTrue3.Value = True Then
        imgRight3.Visible = True
        imgWrong3.Visible = False
    ElseIf optFalse2.Value = True Then
        imgWrong3.Visible = True
        imgRight3.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Randomize
    question(1) = "True or False: A transistor is a moving switch"
    question(2) = "True or False: Moore's Law states that the amount of transistors doubles on a chip every 12 to 18 months"
    question(3) = "True or False: A CPU is like the computers 'brain'"
    question(4) = "Q4"
    question(5) = "Q5"
    question(6) = "Q6"
    question(7) = "Q7"
    question(8) = "Q8"
    question(9) = "Q9"
    question(10) = "Q10"
    question(11) = "Q11"
    question(12) = "Q12"
    question(13) = "Q13"
    question(14) = "Q14"
    question(15) = "Q15"
    question(16) = "Q16"
    question(17) = "Q17"
    question(18) = "Q18"
    question(19) = "Q19"
    question(20) = "Q20"
    
    answer(1) = False
    answer(2) = True
    answer(3) = True
    answer(4) = False
    answer(5) = True
    answer(6) = False
    answer(7) = False
    answer(8) = False
    answer(9) = True
    answer(10) = True
    answer(11) = True
    answer(12) = False
    answer(13) = True
    answer(14) = False
    answer(15) = True
    answer(16) = True
    answer(17) = True
    answer(18) = False
    answer(19) = False
    answer(20) = False
    
    ' Need to set it to value > number of questions so that sub generateQuestions
    ' will randomize the questions
    currentQuestionStart = 100
    generateQuestions
End Sub

