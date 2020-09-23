VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Michael the AIChatter - Alpha"
   ClientHeight    =   6705
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   7050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6120
      Top             =   5730
   End
   Begin VB.TextBox txtRebuild 
      Height          =   345
      Left            =   30
      TabIndex        =   16
      Top             =   8220
      Width           =   7635
   End
   Begin MSComctlLib.ProgressBar Pb1 
      Height          =   225
      Left            =   30
      TabIndex        =   14
      Top             =   6480
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5010
      Top             =   5430
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fmBorderBtm 
      Height          =   135
      Left            =   0
      TabIndex        =   4
      Top             =   5940
      Width           =   9285
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   5745
      Left            =   0
      TabIndex        =   3
      Top             =   210
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   10134
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   285
      Left            =   5790
      TabIndex        =   2
      Top             =   6150
      Width           =   1245
   End
   Begin VB.TextBox txtSend 
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   6120
      Width           =   5715
   End
   Begin VB.Frame fmBorderTop 
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9285
   End
   Begin VB.Frame fmTools 
      Caption         =   "Controls"
      Height          =   5715
      Left            =   5730
      TabIndex        =   5
      Top             =   120
      Width           =   1305
      Begin VB.CommandButton cmdControl 
         Caption         =   "&Shutdown"
         Height          =   350
         Index           =   8
         Left            =   150
         TabIndex        =   17
         ToolTipText     =   "Shutdown and exit AIChatter"
         Top             =   3630
         Width           =   1000
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "&Developer"
         Height          =   350
         Index           =   7
         Left            =   150
         TabIndex        =   13
         ToolTipText     =   "Converts Michael's database into IF/THEN Statements for Sub AIChat."
         Top             =   3210
         Width           =   1000
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "&Clear"
         Height          =   350
         Index           =   6
         Left            =   150
         TabIndex        =   12
         ToolTipText     =   "Clear the Chat Room"
         Top             =   2790
         Width           =   1000
      End
      Begin VB.CommandButton cmdControl 
         Height          =   350
         Index           =   5
         Left            =   150
         TabIndex        =   11
         ToolTipText     =   "Toggles Michael's learning on/off.  Michael's Learning is off"
         Top             =   2370
         Width           =   1000
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Chat &Log"
         Height          =   350
         Index           =   4
         Left            =   150
         TabIndex        =   10
         ToolTipText     =   "View your chat log with Michael"
         Top             =   1950
         Width           =   1000
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "View &NR's"
         Height          =   350
         Index           =   3
         Left            =   150
         TabIndex        =   9
         ToolTipText     =   "View questions/statements that haven't been responded to yet."
         Top             =   1530
         Width           =   1000
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "&View All "
         Height          =   350
         Index           =   2
         Left            =   150
         TabIndex        =   8
         ToolTipText     =   "View all the custom questions/statements Michael has available."
         Top             =   1110
         Width           =   1000
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "&Import"
         Height          =   350
         Index           =   1
         Left            =   150
         TabIndex        =   7
         ToolTipText     =   "Import Database from previous version of AIChatter"
         Top             =   690
         Width           =   1000
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "&Add"
         Height          =   350
         Index           =   0
         Left            =   150
         TabIndex        =   6
         ToolTipText     =   "Add new questions/statements and responses to the database"
         Top             =   270
         Width           =   1000
      End
      Begin VB.Label lblQuestion 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   60
         TabIndex        =   15
         Top             =   4050
         Width           =   1155
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileCreateDB 
         Caption         =   "Create &DB"
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "&Import Previous AIChatter Question File"
      End
      Begin VB.Menu mnuFileView 
         Caption         =   "&View All Questions/Statements in Database"
      End
      Begin VB.Menu mnuFileNR 
         Caption         =   "View &Non-Responded Ques/Stats"
      End
      Begin VB.Menu mnuFileViewChatLog 
         Caption         =   "View Your Chat &Log With Michael"
      End
      Begin VB.Menu mnuFileTglLearn 
         Caption         =   "&Toggle Michael's Learning On/Off"
      End
      Begin VB.Menu mnuFileClear 
         Caption         =   "&Clear ChatBox"
      End
      Begin VB.Menu mnuFileBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&ShutDown"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpStart 
         Caption         =   "&Getting Started"
      End
      Begin VB.Menu mnuHelpBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Michael The AIChatter"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdControl_Click(Index As Integer)
Select Case Index
    Case 0
        ' Create a new general.dat file
        frmAddQues.Show
        Me.Enabled = False
    Case 1
        ' Find a previous AIChatter Version DB
        OpenFile Me
    Case 2
        ' List all questions/statements
        Unload frmList
        FlagStats = "All"
        frmList.Show
    Case 3
        ' List all non responded questions/statements
        Unload frmList
        FlagStats = "NR"
        frmList.Show
    Case 4
        ' View your chat log
        frmView.Show
    Case 5
        ' Toggle Michael's learning on/off
        Call AIToggle
    Case 6
        ' Clear the chat box
        Print #100, frmMain.txtChat.Text
        frmMain.txtChat.Text = ""
    Case 7
        ' Will create IF Then statements for Developer to insert into Public Sub AIChat()
        ' Not implimented at this time
        
        Dim strAns As String
        strAns = InputBox("This tool is for those that have the source code to michael." & _
        Chr(10) & "You must provide a password to continue", "Enter Developer Password")
        If strAns = "" Then Exit Sub
        If strAns = "0wui8l1d2cat" Then
        strAns = MsgBox("This will take your current data file and convert it into" & Chr(10) _
        & "IF/THEN statement for Sub AIChat. This function is only good if you have" & Chr(10) _
        & "The source code for Michael. Do you wish to continue?", vbYesNo, "Convert Data to Subs")
        Else
            Exit Sub
        End If
        Select Case strAns
            Case vbNo
            ' Do nothing
            Case vbYes
               'Dim strSpace As String
                strSpace = "          "
                Open DBFlag For Input As #99
                Open App.Path & "\sys\ifthen.txt" For Append As #98
                    Do While Not EOF(99)
                        Input #99, strQuestion, strAnswer
                        'ElseIf strQuestion = "" Then
                        '    strAnswer = "You didn't say anything."
                        '    AnsFound = True
                        Print #98, "ElseIF instr(strQuestion, " & Chr(34) & strQuestion & Chr(34) & ") Then"
                        Print #98, "strAnswer = " & Chr(34) & strAnswer & Chr(34)
                        Print #98, "AnsFound = True"
                        Print #98,
                    Loop
                MsgBox "IF/THEN Statements Complete.  Your file was saved as " & App.Path & "\sys\ifthen.txt", vbOKOnly, "Complete..."
                Close #99
                Close #98
        End Select
    Case 8
        ' Shutdown the program
        Call ShutDown
End Select
End Sub

Private Sub cmdSend_Click()
frmDebug.List1.Clear
cmdSplit 'txtSend.Text = Split("I, am, a, loser.", ",")
Call AIChat
End Sub

Private Sub Form_Load()

' center the form
frmCenter Me
' Set up the borders
fmBorderTop.Width = Me.Width + 100
fmBorderTop.Left = -50
fmBorderBtm.Width = Me.Width + 100
fmBorderBtm.Left = -50

End Sub

Private Sub mnuFileClear_Click()
cmdControl_Click (6)
End Sub

Private Sub mnuFileCreateDB_Click()
cmdControl_Click (0)
End Sub

Private Sub mnuFileExit_Click()
Call ShutDown
End Sub


Private Sub mnuFileImport_Click()
OpenFile Me
End Sub

Private Sub mnuFileNR_Click()
cmdControl_Click (3)
End Sub

Private Sub mnuFileTglLearn_Click()
cmdControl_Click (5)
End Sub

Private Sub mnuFileView_Click()
cmdControl_Click (2)
End Sub

Private Sub mnuFileViewChatLog_Click()
cmdControl_Click (4)
End Sub

Private Sub mnuHelpAbout_Click()
MsgBox "Michael The AIChatter - Alpha" & Chr(10) _
& "Copyright April 2000   " & Chr(10) & Chr(10) _
& "Creator:  Michael Heath   " & Chr(10) _
& "Email:  mheath@indy.net   " & Chr(10) & Chr(10) _
& "Simulated Chatter. Designed for custom use. Michael's engine allows" & Chr(10) _
& "him to prompt the user for input when he can't answer a question.", vbOKOnly + vbInformation, "About Michael The AIChatter"
End Sub

Private Sub mnuHelpStart_Click()
MsgBox "There is no help availble at the moment." & Chr(10) _
& "Just type a question or statement in the Message Line and send" & Chr(10) _
& "it to the chatroom and see if Michael can return a response. If he" & Chr(10) _
& "can't then he will prompt you to teach him a response as long as his" & Chr(10) _
& "learning ability is on.", vbOKOnly + vbInformation, "Getting Started"
End Sub

Private Sub tmrClose_Timer()
MsgBox "I told you I didn't want to talk to you anymore.", vbOKOnly, "Racist Pig"
ShutDown
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
' Grabs the last sentence sent to michael
If KeyCode = vbKeyUp Then
txtSend.Text = strQuestion
End If

' Sends the sentence to michael
If KeyCode = vbKeyReturn Then
    cmdSend_Click
End If

End Sub

