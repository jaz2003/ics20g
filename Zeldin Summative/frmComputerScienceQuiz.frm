VERSION 5.00
Begin VB.Form frmComputerScienceQuiz 
   BackColor       =   &H00800000&
   Caption         =   "Computer Science Quiz"
   ClientHeight    =   10200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   ScaleHeight     =   10200
   ScaleWidth      =   12315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNewQuestions 
      Caption         =   "New Questions:"
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
      TabIndex        =   19
      Top             =   720
      Width           =   2175
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
      Left            =   8520
      TabIndex        =   2
      Top             =   9000
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
   Begin VB.Label lblQuestion5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   21
      Top             =   6840
      Width           =   3135
   End
   Begin VB.Label lblQuestion4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   20
      Top             =   5520
      Width           =   3135
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
      Left            =   7200
      Picture         =   "frmComputerScienceQuiz.frx":099C
      Top             =   3360
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   17
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label lblQuestion3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   16
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Label lblQuestion1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   15
      Top             =   1800
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
' Number of all predefined questions and answers
Private Const nQuestionsTotal As Integer = 20

' Number of questions in the form presented to the user
Private Const nChoices As Integer = 5

Dim question(1 To nQuestionsTotal) As String
Dim answer(1 To nQuestionsTotal) As Boolean
Dim guess(1 To nChoices) As Boolean

' Index of first of the current questions in question array
Dim currentQuestionStart As Integer

' Generate array of numbers from 1 to NumItems in random order
' Inspired by http://www.vb-helper.com/howto_randomize_array.html
Private Sub randomizeQuestions()
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim tmpQ As String
    Dim tmpA As Boolean
    n = UBound(question)
    
    ' shuffle questions and answers
    For i = 1 To n - 1
        ' Pick a random entry.
        j = Int((n - i + 1) * Rnd + i)
        ' Swap the questions.
        tmpQ = question(i)
        question(i) = question(j)
        question(j) = tmpQ
        ' Swap answers to match the questions.
        tmpA = answer(i)
        answer(i) = answer(j)
        answer(j) = tmpA
    Next i
End Sub
Private Sub generateQuestions()
    ' Pick next group of questions to put in the form
    currentQuestionStart = currentQuestionStart + nChoices
    If currentQuestionStart > nQuestionsTotal Then
        randomizeQuestions
        currentQuestionStart = 1
    End If
    
    lblQuestion1.Caption = question(currentQuestionStart)
    lblQuestion2.Caption = question(currentQuestionStart + 1)
    lblQuestion3.Caption = question(currentQuestionStart + 2)
    lblQuestion4.Caption = question(currentQuestionStart + 3)
    lblQuestion5.Caption = question(currentQuestionStart + 4)
    
End Sub

Private Sub cmdNewQuestions_Click()
    generateQuestions
End Sub

Private Sub cmdReturn4_Click()
    Unload frmComputerScienceQuiz
End Sub

Private Sub cmdSubmit_Click()
    Dim questionAnswered(nChoices) As Boolean
    Dim i As Integer
    
    lblComputerQuizResult.Caption = ""
    
    ' Check whether all questions have been answered
    questionAnswered(1) = optTrue1 Or optFalse1
    questionAnswered(2) = optTrue2 Or optFalse2
    questionAnswered(3) = optTrue3 Or optFalse3
    questionAnswered(4) = True
    questionAnswered(5) = True
    
    For i = 1 To nChoices
        If Not questionAnswered(i) Then
            lblComputerQuizResult.Caption = "Please answer question " + Str(i)
            Exit Sub
        End If
    Next i
    
    ' Collect guesses
    If optTrue1 = answer(currentQuestionStart) Then
        guess(1) = True
        imgRight1.Visible = True
        imgWrong1.Visible = False
    Else
        guess(1) = False
        imgRight1.Visible = False
        imgWrong1.Visible = True
    End If
      
    If optTrue2 = answer(currentQuestionStart + 1) Then
        guess(2) = True
        imgRight2.Visible = True
        imgWrong2.Visible = False
    Else
        guess(2) = False
        imgRight2.Visible = False
        imgWrong2.Visible = True
    End If
    
    If optTrue3 = answer(currentQuestionStart + 2) Then
        guess(3) = True
        imgRight3.Visible = True
        imgWrong3.Visible = False
    Else
        guess(3) = False
        imgRight3.Visible = False
        imgWrong3.Visible = True
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Randomize
    question(1) = "True or False: A transistor is a moving switch"
    question(2) = "True or False: Moore's Law states that the amount of transistors doubles on a chip every 12 to 18 months"
    question(3) = "True or False: A CPU is like the computers 'brain'"
    question(4) = "True or False: The Analytical Engine was an actual computer"
    question(5) = "True or False: RAM stands for Random Access Memory"
    question(6) = "True or False: Apple was founded by Steve Jobs and Steve Wozniak"
    question(7) = "True or False: Computers these days use a DOS interface."
    question(8) = "True or False: GUI stands for Graphicial User Interactions"
    question(9) = "True or False: The LISA computer was a catastrophic fail"
    question(10) = "True or False: Steve Jobs was kicked out of Apple by it's board of directors in 1985."
    question(11) = "True or False: Microsoft was founded by Bill Gates and Paul Allen"
    question(12) = "True or False: The MacOS has the most users"
    question(13) = "True or False: An XOR gate lets current flow through it if both inputs are different"
    question(14) = "True or False: The kernel is the top layer of an OS's structure"
    question(15) = "True or False: A spelling error in your program is called a syntax error"
    question(16) = "True or False: Cobol was one of the first computer languages"
    question(17) = "True or False: Javascript is crucial to the internet"
    question(18) = "True or False: In C++, controls are already programmed for you"
    question(19) = "True or False: In VB 6.0, you need Randomize Timer to generate a random number"
    question(20) = "True or False: A function does not return a value"
    
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

