VERSION 5.00
Begin VB.Form frmComputerScienceQuiz 
   BackColor       =   &H00800000&
   Caption         =   "Computer Science Quiz"
   ClientHeight    =   14000
   ClientLeft      =   225
   ClientTop       =   400
   ClientWidth     =   14000
   LinkTopic       =   "Form1"
   ScaleHeight     =   14000
   ScaleWidth      =   14000
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
      Left            =   10320
      TabIndex        =   13
      Top             =   10800
      Width           =   2175
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
      Height          =   750
      Left            =   7000
      TabIndex        =   2
      Top             =   12500
      Width           =   1600
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
      Left            =   10440
      TabIndex        =   0
      Top             =   11880
      Width           =   1455
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
      Height          =   1800
      Left            =   600
      TabIndex        =   9
      Top             =   1200
      Width           =   3200
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
      Height          =   1800
      Left            =   600
      TabIndex        =   11
      Top             =   3400
      Width           =   3200
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
      Height          =   1800
      Left            =   600
      TabIndex        =   10
      Top             =   5600
      Width           =   3200
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
      Height          =   1800
      Left            =   600
      TabIndex        =   14
      Top             =   7800
      Width           =   3200
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
      Height          =   1800
      Left            =   600
      TabIndex        =   15
      Top             =   10000
      Width           =   3200
   End
   Begin VB.Frame fraQuestion1 
      Height          =   1300
      Left            =   4600
      TabIndex        =   3
      Top             =   1450
      Width           =   1600
      Begin VB.OptionButton optTrue1 
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   1200
      End
      Begin VB.OptionButton optFalse1 
         Caption         =   "False"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   5
         Top             =   780
         Width           =   1200
      End
End
   Begin VB.Frame fraQuestion2 
      Height          =   1300
      Left            =   4600
      TabIndex        =   16
      Top             =   3650
      Width           =   1600
      Begin VB.OptionButton optTrue2 
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   18
         Top             =   300
         Width           =   1200
      End
      Begin VB.OptionButton optFalse2 
         Caption         =   "False"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   17
         Top             =   780
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1300
      Left            =   4600
      TabIndex        =   6
      Top             =   5850
      Width           =   1600
      Begin VB.OptionButton optTrue3 
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   1200
      End
      Begin VB.OptionButton optFalse3 
         Caption         =   "False"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   7
         Top             =   780
         Width           =   1200
      End
   End
   Begin VB.Frame fraQuestion4 
      Height          =   1300
      Left            =   4600
      TabIndex        =   19
      Top             =   8050
      Width           =   1600
      Begin VB.OptionButton optTrue4 
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   21
         Top             =   300
         Width           =   1200
      End
      Begin VB.OptionButton optFalse4 
         Caption         =   "False"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   20
         Top             =   780
         Width           =   1200
      End
   End
   Begin VB.Frame fraQuestion5 
      Height          =   1300
      Left            =   4600
      TabIndex        =   22
      Top             =   10250
      Width           =   1600
      Begin VB.OptionButton optTrue5 
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   23
         Top             =   300
         Width           =   1200
      End
	  Begin VB.OptionButton optFalse5 
         Caption         =   "False"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   24
         Top             =   780
         Width           =   1200
      End
   End
   Begin VB.Image imgRight1 
      Height          =   1590
      Left            =   7000
      Picture         =   "frmComputerScienceQuiz.frx":4E60
      Top             =   1305
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image imgWrong1 
      Height          =   1905
      Left            =   7000
      Picture         =   "frmComputerScienceQuiz.frx":359A
      Top             =   1150
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image imgRight2 
      Height          =   1590
      Left            =   7000
      Picture         =   "frmComputerScienceQuiz.frx":46CB
      Top             =   3505
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image imgWrong2 
      Height          =   1905
      Left            =   7000
      Picture         =   "frmComputerScienceQuiz.frx":2BFE
      Top             =   3350
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image imgRight3 
      Height          =   1590
      Left            =   7000
      Picture         =   "frmComputerScienceQuiz.frx":3F36
      Top             =   5705
      Visible         =   0   'False
      Width           =   1815
   End
      Begin VB.Image imgWrong3 
      Height          =   1905
      Left            =   7000
      Picture         =   "frmComputerScienceQuiz.frx":2262
      Top             =   5550
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image imgRight4 
      Height          =   1590
      Left            =   7000
      Picture         =   "frmComputerScienceQuiz.frx":0000
      Top             =   7905
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image imgWrong4 
      Height          =   1905
      Left            =   7000
      Picture         =   "frmComputerScienceQuiz.frx":18C6
      Top             =   7750
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image imgRight5 
      Height          =   1590
      Left            =   7000
      Picture         =   "frmComputerScienceQuiz.frx":0795
      Top             =   10250
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image imgWrong5 
      Height          =   1905
      Left            =   7000
      Picture         =   "frmComputerScienceQuiz.frx":0F2A
      Top             =   9950
      Visible         =   0   'False
      Width           =   1725
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
      Height          =   1250
      Left            =   600
      TabIndex        =   12
      Top             =   12250
      Width           =   5800
   End
   Begin VB.Label lblComputerQuizTitle 
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

'Used for scoring
Dim points As Integer
Dim score As Integer
Dim level As String


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
    
    'Resets pictures
    imgRight1.Visible = False
    imgRight2.Visible = False
    imgRight3.Visible = False
    imgRight4.Visible = False
    imgRight5.Visible = False
    imgWrong1.Visible = False
    imgWrong2.Visible = False
    imgWrong3.Visible = False
    imgWrong4.Visible = False
    imgWrong5.Visible = False
    
    'Resets option buttons
    optTrue1.Value = False
    optTrue2.Value = False
    optTrue3.Value = False
    optTrue4.Value = False
    optTrue5.Value = False
    optFalse1.Value = False
    optFalse2.Value = False
    optFalse3.Value = False
    optFalse4.Value = False
    optFalse5.Value = False
    
    'Resets points, score, #of times pressed and level
    points = 0
    score = 0
    level = ""
    cmdSubmit.Enabled = True
    'Resets result window
    lblComputerQuizResult.Caption = " "
    
End Sub

Private Sub cmdReturn4_Click()
    Unload frmComputerScienceQuiz
End Sub

Private Sub cmdSubmit_Click()
    Dim questionAnswered(nChoices) As Boolean
    Dim i As Integer
    

    
    ' Check whether all questions have been answered
    questionAnswered(1) = optTrue1 Or optFalse1
    questionAnswered(2) = optTrue2 Or optFalse2
    questionAnswered(3) = optTrue3 Or optFalse3
    questionAnswered(4) = optTrue4 Or optFalse4
    questionAnswered(5) = optTrue5 Or optFalse5
    
    For i = 1 To nChoices
        If Not questionAnswered(i) Then
            lblComputerQuizResult.Caption = "Please answer question " + Str(i)
            Exit Sub
        End If
    Next i
    
    ' Collect guesses
    If optTrue1 = answer(currentQuestionStart) Then
        guess(1) = True
        points = points + 1
        imgRight1.Visible = True
        imgWrong1.Visible = False
    Else
        guess(1) = False
        points = points
        imgRight1.Visible = False
        imgWrong1.Visible = True
    End If
      
    If optTrue2 = answer(currentQuestionStart + 1) Then
        points = points + 1
        guess(2) = True
        imgRight2.Visible = True
        imgWrong2.Visible = False
    Else
        points = points
        guess(2) = False
        imgRight2.Visible = False
        imgWrong2.Visible = True
    End If
    
    If optTrue3 = answer(currentQuestionStart + 2) Then
        guess(3) = True
        points = points + 1
        imgRight3.Visible = True
        imgWrong3.Visible = False
    Else
        guess(3) = False
        points = points
        imgRight3.Visible = False
        imgWrong3.Visible = True
    End If
    
    If optTrue4 = answer(currentQuestionStart + 3) Then
        points = points + 1
        guess(4) = True
        imgRight4.Visible = True
        imgWrong4.Visible = False
    Else
        guess(4) = False
        points = points
        imgRight4.Visible = False
        imgWrong4.Visible = True
    End If
    
    If optTrue5 = answer(currentQuestionStart + 4) Then
        points = points + 1
        guess(5) = True
        imgRight5.Visible = True
        imgWrong5.Visible = False
    Else
        guess(5) = False
        points = points
        imgRight5.Visible = False
        imgWrong5.Visible = True
    End If

    'Calculates percentage
    score = (points / 5) * 100
    'Calculates level
    If score < 50 Then
        level = "R"
    ElseIf score >= 50 And score < 60 Then
        level = "1"
    ElseIf score >= 60 And score < 70 Then
        level = "2"
    ElseIf score >= 70 And score < 80 Then
        level = "3"
    ElseIf score >= 80 And score < 100 Then
        level = "4"
    ElseIf score >= 100 Then
        level = "4++"
    End If
    
    'Determines score
    
    lblComputerQuizResult.Caption = "Hello, you have gotten " & points & " out of 5, which is " & score & "%, and your level is " & level
    
    'Enables only one attempt to submit
        cmdSubmit.Enabled = False
    
End Sub



Private Sub Form_Load()
    Randomize
    question(1) = "True or False: A transistor is a moving switch."
    question(2) = "True or False: A kilobyte is 1024 bytes."
    question(3) = "True or False: A CPU is like the computers 'brain'."
    question(4) = "True or False: The Analytical Engine was an actual computer."
    question(5) = "True or False: RAM stands for Random Access Memory."
    question(6) = "True or False: Apple was founded by Steve Jobs and Steve Wozniak."
    question(7) = "True or False: Computers these days use a DOS interface."
    question(8) = "True or False: GUI stands for Graphicial User Interactions."
    question(9) = "True or False: The LISA computer was a catastrophic fail."
    question(10) = "True or False: Steve Jobs was kicked out of Apple by it's board of directors in 1985."
    question(11) = "True or False: Microsoft was founded by Bill Gates and Paul Allen."
    question(12) = "True or False: The MacOS has the most users."
    question(13) = "True or False: An OR gate allows current to flow through it if at least 1 input is true."
    question(14) = "True or False: The kernel is the top layer of an OS's structure."
    question(15) = "True or False: A spelling error in the message that your program generates is called a syntax error."
    question(16) = "True or False: Cobol was one of the first computer languages."
    question(17) = "True or False: Javascript is crucial to the internet."
    question(18) = "True or False: In C++, controls are already programmed for you."
    question(19) = "True or False: In VB 6.0, you need Randomize Timer to generate a random number."
    question(20) = "True or False: A function does not return a value."
    
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
    answer(15) = False
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

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub



