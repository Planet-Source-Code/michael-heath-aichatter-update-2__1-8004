Attribute VB_Name = "modAIChat"
Public strQuestion As String ' Statement/Question of user
Public strResponse As String ' Comparing string to strQuestion
Public strAnswer As String ' Answer String
Public strName As String ' Name of AIChatter
Public strUser As String ' Name of current user
Public strChkUser As String ' Compare to strUser to see if Michael knows them
Public AnsFound As Boolean
Public AILearn As Boolean   ' Toggle Michaels Learning on or off
Public Hmany As Long ' Counts the number of questions/statements Michael can answer
Public DBFlag As String ' Chooses which database the response is coming from
Public FlagStats As String
Public intWrdCnt As Integer ' Number of words in a sentence
Public splitArray
Public NoTalk As Integer
Public KeyTemp As String ' Temporary file for Sub Keywords
Public strTemp As String ' Temporary Dirty string

Public Sub Main()
' Check to see if the app is already running. If it is, terminate this instance
If App.PrevInstance Then
    MsgBox "AIChatter is already running.", vbOKOnly, "Startup Error"
    End
End If
' Check to see if Michael wants to talk or not
Randomize Timer
frmDebug.Show
On Error Resume Next
' Create system directories
MkDir App.Path & "\data"
MkDir App.Path & "\logs"
MkDir App.Path & "\sys"
' set location of temp file for Sub Keywords
KeyTemp = App.Path & "\sys\tmp1234.sys"
' initiate files
Open App.Path & "\general.dat" For Append As #2
Close #2
Open App.Path & "\data\general.dat" For Append As #2
Close #2
Open App.Path & "\data\unresp.dat" For Append As #2
Close #2
' On first run, we need to move the general.dat file to the right folder
Dim strAns As String
    strAns = ReadINI("INIT", "General", App.Path & "\options.ini")
        If strAns = "Moved" Then
            ' do nothing
        Else
            Open App.Path & "\general.dat" For Input As #2
            Open App.Path & "\data\general.dat" For Append As #3
                Do While Not EOF(2)
                    Input #2, strQuestion, strAnswer
                    Write #3, strQuestion, strAnswer
                Loop
            Close #2
            Close #3
            writeINI "INIT", "General", "Moved", App.Path & "\options.ini"
        End If

' Set the initial database as the general.dat
DBFlag = App.Path & "\data\general.dat"
' Start the system log file
Open App.Path & "\logs\sys.log" For Append As #1
    Print #1, "Give me life - " & Now
strName = "Michael"
' Open a chat log to save the conversation to
Open App.Path & "\logs\chat" & strUser & ".log" For Append As #100
frmMain.Show
frmMain.txtChat.Text = "System:> " & strUser & " has entered the session at " & Now & Chr(10)
frmMain.txtChat.Text = frmMain.txtChat.Text & "System:> " & strName & " has entered the session at " & Now & Chr(10)
Intro
HowMnyQs
AILearn = False ' By default we don't want michael to prompt user for learning
frmMain.cmdControl(5).Caption = "Learn &On"
InitChat
End Sub
Public Sub AIChat()
'On Error GoTo OpenError
' First section for hard coded questions/statements
strQuestion = LCase(frmMain.txtSend.Text)
frmMain.txtSend.Text = ""
If strQuestion = "how are you?" Then
    strAnswer = "i'm fine, how are you?"
    AnsFound = True
ElseIf strQuestion = "" Then
    strAnswer = "You didn't say anything."
    AnsFound = True
ElseIf strQuestion = "<commands>" Then
    strAnswer = Chr(10) & "Here is a list of the current commands:" & Chr(10) & Chr(10) _
    & "<clear>         ----------     Clears the the chat screen" & Chr(10) _
    & "<commands>      ----------     Displays the command list" & Chr(10) _
    & "<delete db>     ----------     Deletes Michael's Database(Will Backup First)" & Chr(10) _
    & "<ai off>        ----------     Turns Michael's Learning Prompter Off" & Chr(10) _
    & "<ai on>         ----------     Turns Michael's Learning Prompter On" & Chr(10) _
    & "<shutdown>      ----------     Shutdown and exit program" & Chr(10) _
    & "<chatlog>       ----------     Displays your current chat log" & Chr(10) _
    & "<stats>         ----------     Displays the questions/statements Michael has responses to."
    AnsFound = True
ElseIf strQuestion = "<stats>" Then
        Unload frmList
        FlagStats = "All"
        frmList.Show
        AnsFound = True
        
ElseIf strQuestion = "<chatlog>" Then
    frmView.Show
    AnsFound = True
    
ElseIf strQuestion = "<shutdown>" Then
    Call ShutDown
    
ElseIf strQuestion = "<ai off>" Then
    frmMain.cmdControl(5).Caption = "Learn O&ff"
    Call AIToggle
    strAnswer = "I will not prompt you when I don't have a response.  All unknown sentences will be logged.  AI is off"
    AnsFound = True
    
ElseIf strQuestion = "<ai on>" Then
    frmMain.cmdControl(5).Caption = "Learn &On"
    Call AIToggle
    strAnswer = "I will prompt you when I don't have a response.  AI is on"
    AnsFound = True
    
ElseIf strQuestion = "<clear>" Then
    frmMain.txtChat.Text = ""
    strAnswer = "I have cleaned the slate for you."
    AnsFound = True
    
ElseIf strQuestion = "<delete db>" Then
    Dim strAns As String
        strAns = MsgBox("This will delete Michael's database and you will have to start over." & Chr(10) _
        & "Are you sure you want to do this?", vbYesNo + vbCritical, "Confirm Deletion")
        Select Case strAns
            Case vbNo
                ' Ok, do nothing but get the hell out of the sub
                Exit Sub
            Case vbYes
                ' make a backup of general.dat as general.bak then delete general.dat
                Open App.Path & "\data\general.dat" For Input As #199
                Open App.Path & "\data\general.bak" For Output As #198
                    Do While Not EOF(199)
                    Input #199, strQuestion, strAnswer
                    Write #198, strQuestion, strAnswer
                    Loop
                Close #198
                Close #199
                Kill App.Path & "\data\general.dat"
                ' Create an empty general.dat
                Open App.Path & "\data\general.dat" For Append As #199
                Close #199
                strQuestion = "<delete db>"
                strAnswer = "You have successfully deleted my brains, I am now stupid again"
                AnsFound = True
                HowMnyQs
        End Select


End If
CustomElseIF
' If a hard coded question was found, then we won't even bother opening
' the questions file
If AnsFound = True Then GoTo SendChat
' Hard coded question not found, so let's search through our general.dat file
Open DBFlag For Input As #2
    Do While Not EOF(2)
    Input #2, strResponse, strAnswer
        LCase (strResponse)
    If strResponse = strQuestion Then
        Print #1, "I have found an answer to the question/statement." & Now
        frmMain.txtChat.Text = frmMain.txtChat.Text & Chr(10) & strUser & ":> " & strQuestion
        frmMain.txtChat.Text = frmMain.txtChat.Text & Chr(10) & strName & ":> " & strAnswer
        AnsFound = True
    End If
    Loop
    Close #2
If AnsFound = True Then GoTo EndChat
KeyWords
If AnsFound = True Then GoTo SendChat
    ' if no response is found, we will take the split sentence and look word by word to
    ' rebuild the sentence into something Michael knows.  If this fails, then
    ' Michael will prompt the user for a saved response.
    ' If AnsFound = False Then Call ReBuildSentence
    
    ' If no answer was found, then we will prompt for a response and then save it to our
    ' general.dat file
    If AnsFound = False Then
        Dim intAns As Integer
        frmMain.txtChat.Text = frmMain.txtChat.Text & Chr(10) & strUser & ":> " & strQuestion
        intAns = Int(Rnd * 5)
        Select Case intAns
            Case 0
                strAnswer = "I guess I don't know what you said, could you rephrase it?"
            Case 1
                strAnswer = "I'm sorry, I'm not the most intelligent so you will have to help me with your last statement"
            Case 2
                strAnswer = "Well, there goes my day. I don't have the slightest idea what you said."
            Case 3
                strAnswer = "uhhh, I must be stoned, try to say that another way."
            Case 4
                strAnswer = "Wow, I'm not doing so well, you'll have to try that again."
            Case 5
                strAnswer = "Damn, I don't have a clue."
        End Select
        'MsgBox intAns
        frmMain.txtChat.Text = frmMain.txtChat.Text & Chr(10) & strName & ":> " & strAnswer
        If AILearn = True Then
            Dim strAdd As String
                strAdd = MsgBox("Would you like to add this question/statement to the database?" & Chr(10) _
                & "'Yes to add, No to place it in the unresponded file, Cancel to disregard.'", vbYesNoCancel, "Unknown Sentence")
                Select Case strAdd
                    Case vbYes
                        ' Call the frmAddQues to add the question/statement
                        frmAddQues.Show
                        frmAddQues.txtQuestion(0).Text = strQuestion
                    Case vbNo
                        ' Save the question to the unresponded file
                        Open App.Path & "\data\unresp.dat" For Append As #2
                            Write #2, strQuestion
                        Close #2
                    Case vbCancel
                        ' do nothing
                End Select
         Else
               ' Save the question to the unresponded file
               Open App.Path & "\data\unresp.dat" For Append As #2
                     Write #2, strQuestion
               Close #2
        End If
                
    End If
GoTo EndChat
OpenError:
    ' No database is found
    strAnswer = "I have encountered a problem. I will attempt to recover from the error.  If the problem persist then please shutdown this program and start it again."
    Print #1, "An error was encountered whiling running Sub AIChat. The error was: " & Err.Description

SendChat:
frmMain.txtChat.Text = frmMain.txtChat.Text & Chr(10) & strUser & ":> " & strQuestion
frmMain.txtChat.Text = frmMain.txtChat.Text & Chr(10) & strName & ":> " & strAnswer
EndChat:
    ' Reset the strings to keep memory clear
    'strQuestion = ""
    strAnswer = ""
    AnsFound = False
frmMain.txtChat.Refresh

Print #1, "AIChat routine " & Now
End Sub
Public Sub HowMnyQs()
' Open the general.dat file and count how many questions are available
Hmany = 0
Open DBFlag For Input As #2
    Do While Not EOF(2)
    Input #2, strQuestion, strAnswer
    Hmany = Hmany + 1
    Loop
Close #2
frmMain.lblQuestion.Caption = "I have " & Hmany & " questions/statements available in my database."
End Sub
Public Sub ShutDown()
On Error Resume Next
Print #100, frmMain.txtChat.Text
Print #1, "I will sleep but not eternally! " & Now
Close All
' delete the temp files
Kill KeyTemp
End
End Sub
Public Sub frmCenter(vForm As Form)
' center the form
vForm.Top = (Screen.Height - vForm.Height) / 2
vForm.Left = (Screen.Width - vForm.Width) / 2

End Sub

Public Sub cmdSplit()
' This example uses the space as the delimiter.
' By doing this you could easily seperate a sentence
' into the individual words it consists of.

splitArray = Split(frmMain.txtSend.Text, " ")
' Here we assign the seperate parts of txtText
' to the array: splitArray.
' By using the space as the delimiter it seperates
' the string at every space.
' Try using different delimiters and see what happens.

For i = 0 To UBound(splitArray)
' Now we will lstbox each string in the array.

frmDebug.List1.AddItem splitArray(i)
Next i
intWrdCnt = i
End Sub

Public Sub InitChat()
' Check to see if we know this user
strChkUser = ReadINI("CHATTERS", strUser, App.Path & "\data\system.dat")
    If strChkUser = "yes" Then
        ' Do something damnit
    End If

End Sub
Public Sub ReBuildSentence()
' Here we will try to rebuild the sentence to something Michael understands and then re-compare it
' against the database
' Not yet implimented
Dim i As Integer
    For i = 0 To intWrdCnt
        frmMain.txtRebuild.Text = frmMain.txtRebuild.Text & splitArray(i) & " "
        ' Do the word search and replace as needed
    Next i
End Sub
Public Sub AIToggle()
        If frmMain.cmdControl(5).Caption = "Learn O&ff" Then
            frmMain.cmdControl(5).Caption = "Learn &On"
            AILearn = False
            frmMain.cmdControl(5).ToolTipText = "Toggles Michael's learning on/off.  Michael's Learning is off"
        ElseIf frmMain.cmdControl(5).Caption = "Learn &On" Then
            frmMain.cmdControl(5).Caption = "Learn O&ff"
            AILearn = True
            frmMain.cmdControl(5).ToolTipText = "Toggles Michael's learning on/off.  Michael's Learning is on"
        End If

End Sub
Public Sub Intro()
strAnswer = "Hello " & strUser & "!" & Chr(10) _
& "Nice to see you make it here.  I'm not sure what I know at the moment" & Chr(10) _
& "Just start typing and we'll see if I can respond.  If I don't have a response," & Chr(10) _
& "then I will prompt you to help me out as long as my AI Learning is on." & Chr(10) _
& Chr(10) & "Type <commands> to see a list of the current commands available." & Chr(10) _
& "Above all, if you have the patience, I can learn a lot from you."
frmMain.txtChat.Text = frmMain.txtChat.Text & strName & ":> " & strAnswer

End Sub
Public Sub LogQuestion()
Open App.Path & "\data\unresp.dat" For Append As #99
    Write #99, strQuestion
Close #99
End Sub
