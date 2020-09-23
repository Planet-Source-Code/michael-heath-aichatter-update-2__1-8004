VERSION 5.00
Begin VB.Form frmAddQues 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create New Questions/Statements and Responses"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAnswer 
      Height          =   975
      Left            =   750
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3120
      Width           =   5325
   End
   Begin VB.TextBox txtQuestion 
      Height          =   315
      Index           =   3
      Left            =   750
      TabIndex        =   5
      Top             =   2100
      Width           =   5325
   End
   Begin VB.TextBox txtQuestion 
      Height          =   315
      Index           =   2
      Left            =   750
      TabIndex        =   4
      Top             =   1590
      Width           =   5325
   End
   Begin VB.TextBox txtQuestion 
      Height          =   315
      Index           =   1
      Left            =   750
      TabIndex        =   3
      Top             =   1080
      Width           =   5325
   End
   Begin VB.TextBox txtQuestion 
      Height          =   285
      Index           =   0
      Left            =   780
      TabIndex        =   0
      Top             =   360
      Width           =   5325
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Accept"
      Height          =   405
      Index           =   0
      Left            =   3570
      TabIndex        =   12
      Top             =   4200
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   405
      Index           =   1
      Left            =   4830
      TabIndex        =   13
      Top             =   4200
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Type the response you wish Michael to make for the previous question/statement:"
      Height          =   405
      Index           =   2
      Left            =   750
      TabIndex        =   11
      Top             =   2520
      Width           =   5325
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Original:"
      Height          =   255
      Index           =   3
      Left            =   30
      TabIndex        =   9
      Top             =   390
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3."
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   2130
      Width           =   345
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2."
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   1620
      Width           =   345
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1."
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   1110
      Width           =   345
   End
   Begin VB.Label Label1 
      Caption         =   "Please type any variations of the question/statement here(Optional but recommended):"
      Height          =   465
      Index           =   1
      Left            =   780
      TabIndex        =   2
      Top             =   660
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Please type your original question/statement here:"
      Height          =   195
      Index           =   0
      Left            =   780
      TabIndex        =   1
      Top             =   120
      Width           =   4725
   End
End
Attribute VB_Name = "frmAddQues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
        Open App.Path & "\data\general.dat" For Append As #2
        Close #2
        If txtQuestion(0).Text = "" Then
            MsgBox "You must provide a question/statement in the first text box to continue.", vbOKOnly, "Error"
            Exit Sub
        End If
        If txtAnswer.Text = "" Then
            MsgBox "You must provide a response for your question/statement to continue.", vbOKOnly, "Error"
            Exit Sub
        End If
        ' Save the new questions/statements and response
            ' let's make sure none of the new statements are in the database first
        Open App.Path & "\data\general.dat" For Input As #2
            Do While Not EOF(2)
            Input #2, strQuestion, strAnswer
            If LCase(txtQuestion(0).Text) = strQuestion Then
                txtQuestion(0).Text = ""
            End If
            If LCase(txtQuestion(1).Text) = strQuestion Then
                txtQuestion(1).Text = ""
            End If
            If LCase(txtQuestion(2).Text) = strQuestion Then
                txtQuestion(2).Text = ""
            End If
            If LCase(txtQuestion(3).Text) = strQuestion Then
                txtQuestion(3).Text = ""
            End If
        Loop
        Close #2
            
        Open App.Path & "\data\general.dat" For Append As #2
            Write #2, LCase(txtQuestion(0).Text), txtAnswer.Text
            If txtQuestion(1).Text > "" Then
                Write #2, LCase(txtQuestion(1).Text), txtAnswer.Text
            End If
            If txtQuestion(2).Text > "" Then
                Write #2, LCase(txtQuestion(2).Text), txtAnswer.Text
            End If
            If txtQuestion(3).Text > "" Then
                Write #2, LCase(txtQuestion(3).Text), txtAnswer.Text
            End If
        Close #2
    Case 1
        ' Cancel
End Select
        frmMain.Enabled = True
        Unload Me
        frmMain.Show
        HowMnyQs
End Sub

Private Sub Form_Load()
frmCenter Me
frmMain.Enabled = False
End Sub
