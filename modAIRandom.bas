Attribute VB_Name = "modAIRandom"
Public Sub AIRandom()
Dim intI As Integer
intI = Int(Rnd * 21)
Select Case intI
    Case 0
        strAnswer = "I don't have anymore to say about that, " & strUser
    Case 1
        strAnswer = strUser & ", really, could we just talk about something else?"
    Case 2
        strAnswer = "You have a one track mind don't you?"
    Case 3
        strAnswer = "I'm getting very bored. Change the freakin subject!"
    Case 4
        strAnswer = "Ok, I've just about had enough of this."
    Case 5
        strAnswer = "Wow, don't you ever give up?"
    Case 6
        strAnswer = "How much longer will you annoy me with this garbage?"
    Case 7
        strAnswer = "I don't want to talk about that anymore."
    Case 8
        strAnswer = "Do you really think I have anything else to say about it?"
    Case 9
        strAnswer = "You are a persistent little person"
    Case 10
        strAnswer = "Is there any reason for this nonsense?"
    Case 11
        strAnswer = "Why must you continue to drive me crazy with it???"
    Case 12
        strAnswer = "I'm beggin you to talk about something else"
    Case 13
        strAnswer = "Ok, if you can't think of anything else to talk about, I'm just going to log you off"
    Case 14
        strAnswer = "Just keep it up."
    Case 15
        strAnswer = "I'm not kidding!!!"
    Case 16
        strAnswer = "I'm getting really sick of you and this dialog."
    Case 17
        strAnswer = " :("
    Case 18
        strAnswer = " - "
    Case 19
        strAnswer = "I've got better things to do"
    Case 20
        strAnswer = "You're last chance is coming up!"
    Case 21
        strAnswer = "Take a hike pal!"
End Select
End Sub

