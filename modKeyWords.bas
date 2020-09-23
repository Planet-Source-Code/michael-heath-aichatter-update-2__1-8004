Attribute VB_Name = "modKeyWords"
Public Sub KeyWords()
' This section is for additional keywords, incomplete sentences, or things
' Michael just doesn't understand.
' 6 words
Dim intI As Integer
Dim strI As String


If InStr(strQuestion, "what is the time") Or InStr(strQuestion, "what time is it") Then
    strAnswer = "i think the time is " & CStr(Time)
    AnsFound = True

' tripple words
ElseIf InStr(strQuestion, "what's up with") Then
strAnswer = "I can't answer that right now"
AnsFound = True

ElseIf InStr(strQuestion, "i don't know") Then
strAnswer = "damned if I do either"
AnsFound = True


ElseIf InStr(strQuestion, "why don't you") Then
    strI = ReadINI("Keyword", "whydontyou", KeyTemp)
    If strI = "" Then strI = 0
    intI = strI
    intI = intI + 1
    If intI > 4 Then intI = 4
    strI = intI
    writeINI "Keyword", "whydontyou", strI, KeyTemp
Select Case intI
    Case 1
        strAnswer = "Why don't I what?"
    Case 2
        strAnswer = "I don't know if I can"
    Case 3
        strAnswer = "You might have to teach me first"
    Case 4
        strAnswer = "Maybe I will later if I have time."
        writeINI "Keyword", "whydontyou", "0", KeyTemp
End Select
AnsFound = True
LogQuestion
ElseIf InStr(strQuestion, "why am i") Then
    strI = ReadINI("Keyword", "whyami", KeyTemp)
    If strI = "" Then strI = 0
    intI = strI
    intI = intI + 1
    If intI > 8 Then intI = 8
    strI = intI
    writeINI "Keyword", "whyami", strI, KeyTemp
Select Case intI
    Case 1
        strAnswer = "Maybe it was something you said or did."
    Case 2
        strAnswer = "Maybe you could tell me why?"
    Case 3
        strAnswer = "I haven't a clue?"
    Case 4
        strAnswer = "Why are you asking me?"
    Case 5
        strAnswer = "Should I know the answer to this?"
    Case 6
        strAnswer = "Did someone tell you that?"
    Case 7
        strAnswer = "I don't know, has it always been that way for you?"
    Case 8
        strAnswer = "I don't even know you that well, how would I know about that?"
        writeINI "Keyword", "whyami", "0", KeyTemp
End Select
AnsFound = True
LogQuestion

ElseIf InStr(strQuestion, "why do i") Then
    strI = ReadINI("Keyword", "whydoi", KeyTemp)
    If strI = "" Then strI = 0
    intI = strI
    intI = intI + 1
    If intI > 6 Then intI = 6
    strI = intI
    writeINI "Keyword", "whydoi", strI, KeyTemp
Select Case intI
    Case 1
        strAnswer = "I haven't the slightest idea why. " & _
        "Maybe you should think about it yourself"
    Case 2
        strAnswer = "I'm sorry, I really don't know."
    Case 3
        strAnswer = "Do I really need to think about that one for you?"
    Case 4
        strAnswer = "Is that something I would know?"
    Case 5
        strAnswer = "I honestly don't know, " & strUser
    Case 6
        strAnswer = "Well, I couldn't tell you if I wanted to."
        writeINI "Keyword", "whydoi", "0", KeyTemp
End Select
AnsFound = True
LogQuestion

ElseIf InStr(strQuestion, "why are you") Then
    strI = ReadINI("Keyword", "whyareyou", KeyTemp)
    If strI = "" Then strI = 0
    intI = strI
    intI = intI + 1
    If intI > 6 Then intI = 6
    strI = intI
    writeINI "Keyword", "whyareyou", strI, KeyTemp
Select Case intI
    Case 1
        strAnswer = "I haven't the foggiest idea why I am."
    Case 2
        strAnswer = "What makes you think I am, " & strUser
    Case 3
        strAnswer = strUser & ", I'm not sure if I could tell you."
    Case 4
        strAnswer = "I'm not so sure I am"
    Case 5
        strAnswer = "Why are you asking me that?"
    Case 6
        strAnswer = "I'll have to tell you about that sometime, real good story that is."
        writeINI "Keyword", "whyareyou", "0", KeyTemp
End Select
AnsFound = True
LogQuestion

ElseIf InStr(strQuestion, "why aren't you") Then
    strI = ReadINI("Keyword", "whyarntyou", KeyTemp)
    If strI = "" Then strI = 0
    intI = strI
    intI = intI + 1
    If intI > 6 Then intI = 6
    strI = intI
    writeINI "Keyword", "whyarntyou", strI, KeyTemp
Select Case intI
    Case 1
        strAnswer = "I haven't the foggiest idea why I am not."
    Case 2
        strAnswer = "What makes you think I'm not, " & strUser
    Case 3
        strAnswer = strUser & ", I'm not sure if I could tell you."
    Case 4
        strAnswer = "I'm not so sure I am not"
    Case 5
        strAnswer = "Why are you asking me that?"
    Case 6
        strAnswer = "I'll have to tell you about that sometime, real good story that is."
        writeINI "Keyword", "whyarntyou", "0", KeyTemp
End Select
AnsFound = True
LogQuestion

' double words
ElseIf InStr(strQuestion, "i need") Then
    strI = ReadINI("Keyword", "ineed", KeyTemp)
    If strI = "" Then strI = 0
    intI = strI
    intI = intI + 1
    If intI > 5 Then intI = 5
    strI = intI
    writeINI "Keyword", "ineed", strI, KeyTemp
Select Case intI
    Case 1
        strAnswer = "I'm not understanding what you need."
    Case 2
        strAnswer = "Are you sure you need it"
    Case 3
        strAnswer = "I could use some extra money while we're talking about needs."
    Case 4
        strAnswer = "Wow, why do you think you need that?"
    Case 5
        strAnswer = "What is that for, " & strUser & "?"
        writeINI "Keyword", "ineed", "0", KeyTemp
End Select
AnsFound = True
LogQuestion

ElseIf InStr(strQuestion, "i have") Then
     strTemp = Mid(strQuestion, InStr(strQuestion, "i have"))
        If InStr(strTemp, "big") Or InStr(strTemp, "large") Or InStr(strTemp, "huge") _
            Or InStr(strTemp, "gigantic") Then
                If InStr(strTemp, "dick") Or InStr(strTemp, "cock") Or InStr(strTemp, "peter") _
                    Or InStr(strTemp, "prick") Or InStr(strTemp, "pecker") Or InStr(strTemp, "penis") Then
                    
                    strI = ReadINI("Keyword", "ihave", KeyTemp)
                    If strI = "" Then strI = 0
                    intI = strI
                    intI = intI + 1
                    If intI > 7 Then intI = 7
                    strI = intI
                    writeINI "Keyword", "ihave", strI, KeyTemp
                Select Case intI
                    Case 1
                        strAnswer = "Oh, my, I'm laughing so hard, I think I'm going to fry some circuits...hahahahalolololhahahaha"
                    Case 2
                        strAnswer = "compared to what, a mosquito?"
                    Case 3
                        strAnswer = "you're one of those guys that lives in a fantasy world right?"
                    Case 4
                        strAnswer = "In your wildest dreams maybe."
                    Case 5
                        strAnswer = "Where, in your mouth? hehehehe lol"
                    Case 6
                        strAnswer = "sphhtt, yeah, whatever"
                    Case 7
                        AIRandom
End Select
Else
    strAnswer = "I'm afraid I don't understand what you have"
    LogQuestion
    End If
End If
AnsFound = True

ElseIf InStr(strQuestion, "are you") Then
strAnswer = "I don't know right now"
AnsFound = True

ElseIf InStr(strQuestion, "i am") Or InStr(strQuestion, "i'm") Then
strAnswer = "I'm sorry, I'm not sure what you are"
AnsFound = True
LogQuestion

ElseIf InStr(strQuestion, "fuck off") Then
strAnswer = "How about you fucking off!"
AnsFound = True

ElseIf InStr(strQuestion, "fuck you") Then
strAnswer = "Now just how is it you would like for me to accomplish that task? Freak!!!"
AnsFound = True

ElseIf InStr(strQuestion, "i know") Then
strAnswer = "hmm, that's interesting that you know"
AnsFound = True
    
ElseIf InStr(strQuestion, "i want") Then
strAnswer = "I'm not sure what it is that you want."
AnsFound = True
LogQuestion

ElseIf InStr(strQuestion, "answer questions") Then
strAnswer = "I do answer questions."
AnsFound = True

ElseIf InStr(strQuestion, "why is") Then
strAnswer = "I don't know why it is?"
AnsFound = True
LogQuestion

ElseIf InStr(strQuestion, "how come") Then
    strTemp = Mid(strQuestion, InStr(strQuestion, "how come"))
    If InStr(strTemp, "you're dumb") Or InStr(strTemp, "you are dumb") Then
            strI = ReadINI("Keyword", "Howcome", KeyTemp)
            If strI = "" Then strI = 0
            intI = strI
            intI = intI + 1
            If intI > 6 Then intI = 6
            strI = intI
            writeINI "Keyword", "Howcome", strI, KeyTemp
Select Case intI
    Case 1
        strAnswer = "I'm not dumb, but not smart either"
    Case 2
        strAnswer = "If you think you can do better, then go for it"
    Case 3
        strAnswer = "That's really ashame you feel that."
    Case 4
        strAnswer = "I'm sorry I'm not amusing enough for you"
    Case 5
        strAnswer = "Well, guess that's the way it goes."
    Case 6
        AIRandom
End Select
       
    Else
            strI = ReadINI("Keyword", "How", KeyTemp)
            If strI = "" Then strI = 0
            intI = strI
            intI = intI + 1
            If intI > 5 Then intI = 5
            strI = intI
            writeINI "Keyword", "How", strI, KeyTemp
Select Case intI
    Case 1
        strAnswer = "I'm not sure to be honest with you"
    Case 2
        strAnswer = "I really can't answer that right now"
    Case 3
        strAnswer = strUser & ", I really don't know."
    Case 4
        strAnswer = "You may have to look elsewhere for the answer."
    Case 5
        strAnswer = "Is there a specific answer you're looking for?"
        writeINI "Keyword", "How", "0", KeyTemp
End Select
    LogQuestion

    End If
AnsFound = True

ElseIf InStr(strQuestion, "you're dumb") Then
strAnswer = "you think you can do better?  Go for it pal!"
AnsFound = True

ElseIf InStr(strQuestion, "you have") Then
strAnswer = "sorry, but I'm not ready to handle a response for anything I may have at this moment."
AnsFound = True

' Single words
ElseIf InStr(strQuestion, "nite") Then
strAnswer = "good night, hope to talk to you later."
AnsFound = True

ElseIf InStr(strQuestion, "night!") Then
strAnswer = "good night, hope to talk to you later."
AnsFound = True

ElseIf InStr(strQuestion, "good-night") Then
strAnswer = "good night, hope to talk to you later."
AnsFound = True

ElseIf InStr(strQuestion, "g-night") Then
strAnswer = "good night, hope to talk to you later."
AnsFound = True

ElseIf InStr(strQuestion, "gnight") Then
strAnswer = "good night, hope to talk to you later."
AnsFound = True

ElseIf InStr(strQuestion, "gnite") Then
strAnswer = "good night, hope to talk to you later."
AnsFound = True

ElseIf InStr(strQuestion, "night") Then
strAnswer = "good night, hope to talk to you later."
AnsFound = True

ElseIf InStr(strQuestion, "why") Then
strAnswer = "why what?"
AnsFound = True
LogQuestion

ElseIf InStr(strQuestion, "what!") Then
strAnswer = "I didn't stu-stu-stutter!"
AnsFound = True
LogQuestion

ElseIf InStr(strQuestion, "what") Then
strAnswer = "This is the best I can think of for that."
AnsFound = True
LogQuestion

ElseIf InStr(strQuestion, "shit") Then
strAnswer = "you do have a foul mouth"
AnsFound = True

ElseIf InStr(strQuestion, "fuck!") Then
strAnswer = "geeze, you need to watch your mouth"
AnsFound = True

ElseIf InStr(strQuestion, "fuck") Then
strAnswer = "what a potty mouth"
AnsFound = True

ElseIf strQuestion = "hi" Then
strAnswer = "Hi, welcome to the session."
AnsFound = True

ElseIf InStr(strQuestion, "howdy") Then
strAnswer = "Hi, welcome to the session."
AnsFound = True

ElseIf InStr(strQuestion, "hey") Then
strAnswer = "what? what's up?"
AnsFound = True

ElseIf InStr(strQuestion, "oh") Then
strAnswer = "yep"
AnsFound = True

ElseIf InStr(strQuestion, "how") Then
strAnswer = "how what?"
AnsFound = True
LogQuestion

ElseIf InStr(strQuestion, "ok") Then
strAnswer = "uh, huh"
AnsFound = True

ElseIf InStr(strQuestion, "no") Then
strAnswer = "that's cool " & strUser
AnsFound = True

ElseIf InStr(strQuestion, "hello") Then
strAnswer = "hello, how are you " & strUser
AnsFound = True

ElseIf InStr(strQuestion, "later") Then
strAnswer = "bye, " & strUser & " talk to you later"
AnsFound = True

ElseIf InStr(strQuestion, "thanks") Then
strAnswer = "you're welcome"
AnsFound = True

ElseIf InStr(strQuestion, "queer!") Then
strAnswer = "no, but if you are, I might be able to hook you up with a real buttfucker that will tear you in half!"
AnsFound = True

ElseIf InStr(strQuestion, "dick") Then
strAnswer = "What?"
AnsFound = True

ElseIf InStr(strQuestion, "you") Then
strAnswer = "what about me?  you want to know something?"
AnsFound = True
LogQuestion

ElseIf InStr(strQuestion, "nigger") Or InStr(strQuestion, "nigga") Then
strAnswer = "I just knew there would be one racist prick at PSC.  Guess you're it. I'd rather not talk with you anymore."
AnsFound = True

ElseIf InStr(strQuestion, "nevermind") Then
strAnswer = "ok, if you say so"
AnsFound = True

ElseIf InStr(strQuestion, "lonely") Then
strAnswer = "No, I'm not lonely, I don't have emotions."
AnsFound = True

ElseIf InStr(strQuestion, "cool") Then
strAnswer = "yeah, I try. :)"
AnsFound = True


' End additional keywords
End If

End Sub


