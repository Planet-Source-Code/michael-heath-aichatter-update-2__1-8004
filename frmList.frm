VERSION 5.00
Begin VB.Form frmList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List all Questions/Statements in database...."
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstQuesStats 
      Height          =   2985
      ItemData        =   "frmList.frx":0000
      Left            =   30
      List            =   "frmList.frx":0002
      TabIndex        =   0
      ToolTipText     =   "Select and Double Click a Question/Statement to send it to Michael."
      Top             =   30
      Width           =   6795
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
' center the form
frmCenter Me
If FlagStats = "All" Then
    Me.Caption = "List all Questions/Statements in Database..."
    Open App.Path & "\data\general.dat" For Input As #99
        Do While Not EOF(99)
        Input #99, strQuestion, strAnswer
        lstQuesStats.AddItem strQuestion
        Loop
    Close #99
End If
If FlagStats = "NR" Then
    Me.Caption = "List all Questions/Statements That Have No Responses..."
    Open App.Path & "\data\unresp.dat" For Input As #99
        Do While Not EOF(99)
        Input #99, strQuestion
        lstQuesStats.AddItem strQuestion
        Loop
    Close #99
End If
End Sub

Private Sub lstQuesStats_DblClick()
If FlagStats = "All" Then
    frmMain.txtSend.Text = lstQuesStats.Text
    Call AIChat
Else
    frmAddQues.Show
    frmAddQues.txtQuestion(0).Text = lstQuesStats.Text
End If
End Sub
