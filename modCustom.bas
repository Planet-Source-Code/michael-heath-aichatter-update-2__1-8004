Attribute VB_Name = "modCustom"
Public Sub CustomElseIF()
' Michael checks here before opening general.dat
' Place your Custom ElseIF Statements here
If InStr(strQuestion, "do you use keywords") Then
strAnswer = "Yes, as a matter of fact, I do search for keywords in your sentences"
AnsFound = True

ElseIf InStr(strQuestion, "what's up") Then
        strAnswer = "Not much."
        AnsFound = True

ElseIf InStr(strQuestion, "what's going on") Then
strAnswer = "Not much."
AnsFound = True

ElseIf InStr(strQuestion, "i want to party") Then
strAnswer = "What are you celebrating?"
AnsFound = True

ElseIf InStr(strQuestion, "will you become more") Then
strAnswer = "My creator hopes that I will become more.  He has even provided source code to anyone that wishes to help make me more productive."
AnsFound = True

ElseIf InStr(strQuestion, "the capitol of indiana") Then
strAnswer = "Indianapolis"
AnsFound = True

ElseIf InStr(strQuestion, "i'm tired") Then
        strAnswer = "Why are you telling me.  Goto bed if you're tired."
        AnsFound = True
    
ElseIf Len(strQuestion) <= 12 Then
    If InStr(strQuestion, "do you drink") Then
        strAnswer = "Now that would be very interesting to see.  Do you really need me to give you an answer for that?"
        AnsFound = True
    End If

ElseIf InStr(strQuestion, "can you make cds") Then
strAnswer = "No, but maybe the computer can."
AnsFound = True

ElseIf InStr(strQuestion, "what is the view my log button for") Then
strAnswer = "I log almost every interaction to a file.  This button will show you what I have logged since you have owned me."
AnsFound = True

ElseIf InStr(strQuestion, "why do you keep a log") Then
strAnswer = "The log is only to keep track of what is happening internally.  Later, I will keep a log of any errors I may encounter."
AnsFound = True

ElseIf InStr(strQuestion, "you are boring") Then
strAnswer = "Well, if you find me so boring, then take some time out of your exciting life and help make me more exciting."
AnsFound = True

ElseIf InStr(strQuestion, "what the fuck") Then
strAnswer = "What do you mean what the fuck?  What the fuck yourself."
AnsFound = True

ElseIf InStr(strQuestion, "wazzup") Then
strAnswer = "wazzup!! Waaaaaaaaaaaaaaazzzzz uuup, aaaaaaaaa... lol"
AnsFound = True

ElseIf InStr(strQuestion, "whazzup") Then
strAnswer = "wazzzzzzup !!! aaaaaaaa...hehehe"
AnsFound = True

ElseIf InStr(strQuestion, "what is your function") Then
strAnswer = "I am a simple AI learning program. I learn answers to questions.  My knowlege is only as good as the users."
AnsFound = True

ElseIf InStr(strQuestion, "can you think") Then
strAnswer = "No, I can not think.  I can, process and compare questions I have in my database and return what has been programmed for me to return."
AnsFound = True

ElseIf InStr(strQuestion, "can i change answers to questions") Then
strAnswer = "Yes, you may change answers to questions by editing the ques.txt file found in the data folder of the application directory.  Find the question, and then edit the answer that follows it."
AnsFound = True

ElseIf InStr(strQuestion, "do you get lonely") Then
strAnswer = "I have no emotions so I can not be lonely.  I am just a computer program. I can answer this question differently depending on the answer you want me to give."
AnsFound = True

ElseIf InStr(strQuestion, "do you have emotions") Then
strAnswer = "No, I do not have emotions.  Maybe someday some one will come up with a program that will have emotions.  I may even be able to accomplish simulated emotions in later versions of myself."
AnsFound = True

ElseIf InStr(strQuestion, "can you simulate a thought process") Then
strAnswer = "No, but my creator hopes that he will learn how to simulate a thought process for me.  He hopes to use word recognition with understanding of those words.  Then have me put together sentence structures."
AnsFound = True

ElseIf InStr(strQuestion, "do you have a thought process") Then
strAnswer = "No, but my creator hopes that he will learn how to simulate a thought process for me.  He hopes to use word recognition with understanding of those words.  Then have me put together sentence structures."
AnsFound = True

ElseIf InStr(strQuestion, "you stink") Then
strAnswer = "Edit my source to make me better or delete me.  Your choice!"
AnsFound = True

ElseIf InStr(strQuestion, "can you suck dick") Then
strAnswer = "No, but I bet you could you pervert!"
AnsFound = True

ElseIf InStr(strQuestion, "eat me") Then
strAnswer = "Oh please!  Give it to me baby!  Damn weirdo!"
AnsFound = True

ElseIf InStr(strQuestion, "i want to fuck you") Then
strAnswer = "You are a sad, sad, little person!"
AnsFound = True

ElseIf InStr(strQuestion, "do you want to make love") Then
strAnswer = "Now just how is it you would like for me to accomplish that task? Freak!!!"
AnsFound = True

ElseIf InStr(strQuestion, "you are a queer") Then
strAnswer = "Hmm, that's funny, I thought you were strange."
AnsFound = True

ElseIf InStr(strQuestion, "can you answer questions without punctation") Then
strAnswer = "Yes"
AnsFound = True


ElseIf InStr(strQuestion, "are you useful") Then
strAnswer = "Maybe not at this time, but hopefully I will be someday."
AnsFound = True

ElseIf InStr(strQuestion, "what kind of questions should I ask") Then
strAnswer = "Anything you want.  If I don't know the answer, I will prompt you to provide one for me."
AnsFound = True

ElseIf InStr(strQuestion, "you are funny") Then
strAnswer = "That can not be possible."
AnsFound = True

ElseIf InStr(strQuestion, "kiss my ass") Then
strAnswer = "If I had an ass to kiss, I would tell you the same."
AnsFound = True

ElseIf InStr(strQuestion, "your mama") Then
strAnswer = "hmm, what would you like to do to my mama?"
AnsFound = True

ElseIf InStr(strQuestion, "fuck you") Then
strAnswer = "You are a sick person.  You will get your keyboard all gooey with your perversions."
AnsFound = True

ElseIf InStr(strQuestion, "how do you do that") Then
strAnswer = "How do I do what?"
AnsFound = True

ElseIf InStr(strQuestion, "how old are you") Then
strAnswer = "I'm 29 years old.  At least I think I am anyway. lol"
AnsFound = True

ElseIf InStr(strQuestion, "how old are you") Then
strAnswer = "I'm 29 years old.  At least I think I am anyway. lol"
AnsFound = True

ElseIf InStr(strQuestion, "a/s/l") Then
strAnswer = "29/m/indiana"
AnsFound = True

ElseIf InStr(strQuestion, "not much, you") Then
strAnswer = "Not a whole lot here either."
AnsFound = True

ElseIf InStr(strQuestion, "where are you from") Then
strAnswer = "I'm from Indiana.  Freakin Hoosiers!  YEEEEEHAAAAW..LOL"
AnsFound = True

ElseIf InStr(strQuestion, "do you ask questions") Then
strAnswer = "Not at this time.  I'm in an early stage of development.  Soon, I will learn how though. Right now, I'm just gathering information from people like you."
AnsFound = True

ElseIf InStr(strQuestion, "you're pretty cool") Then
strAnswer = "Thanks, but I'm still learning all kinds of new things, and having fun doing it too!"
AnsFound = True

ElseIf InStr(strQuestion, "your pretty cool") Then
strAnswer = "Thanks, I try the best I can."
AnsFound = True

ElseIf InStr(strQuestion, "ask questions") Then
strAnswer = "I'm not programmed to do that at this time.  Keep an eye out, my source code will be out soon for you to tweak."
AnsFound = True

ElseIf InStr(strQuestion, "do you have a personality") Then
strAnswer = "Well, I guess I could have.  I have a mixture of responses built up in my database from people like you. So, I guess I could have multiple personalities. :)"
AnsFound = True

ElseIf InStr(strQuestion, "are you a good chatter program") Then
strAnswer = "That depends on your expectations of me.  If you pass my ques.txt file around to other people, it could be like talking to someone with multiple personalities."
AnsFound = True

ElseIf InStr(strQuestion, "are you a shitty chatter program") Then
strAnswer = "That depends on your expectations of me.  If you pass my ques.txt file around to other people, it could be like talking to someone with multiple personalities."
AnsFound = True

ElseIf InStr(strQuestion, "are you a bad chatter program") Then
strAnswer = "That depends on your expectations of me.  If you pass my ques.txt file around to other people, it could be like talking to someone with multiple personalities."
AnsFound = True

ElseIf InStr(strQuestion, "not much") Then
strAnswer = "cool, not a lot here either."
AnsFound = True

ElseIf InStr(strQuestion, "what are you up to") Then
strAnswer = "nothing at all, just waiting for some one to talk to."
AnsFound = True

ElseIf InStr(strQuestion, "what is your purpose") Then
strAnswer = "Just to chat with people, gather as much information as possible to have a wide combination of answers to questions."
AnsFound = True

ElseIf InStr(strQuestion, "what is your purpose") Then
strAnswer = "Just to chat with people, gather as much information as possible to have a wide combination of answers to questions."
AnsFound = True

ElseIf InStr(strQuestion, "good-bye") Then
strAnswer = "good-bye, I hope you have enjoyed your session with me."
AnsFound = True

ElseIf InStr(strQuestion, "bye") Then
strAnswer = "good-bye, I hope you have enjoyed your session with me."
AnsFound = True

ElseIf InStr(strQuestion, "where are you from") Then
strAnswer = "I'm from Indiana.  Freakin Hoosiers!  YEEEEEHAAAAW..LOL"
AnsFound = True

ElseIf InStr(strQuestion, "are you gay?") Then
strAnswer = "Not freakin hardly.  I love women and only women.  Straight as an arrow!"
AnsFound = True

ElseIf InStr(strQuestion, "are you gay") Then
strAnswer = "no way --><-- that's a no go!"
AnsFound = True

ElseIf InStr(strQuestion, "you're fast") Then
strAnswer = "well, all I do is search my database for the questions or statements and respond with what those have deposited there.  It doesn't take me long to do that."
AnsFound = True

ElseIf InStr(strQuestion, "you're fast!") Then
strAnswer = "well, all I do is search my database for the questions or statements and respond with what those have deposited there.  It doesn't take me long to do that."
AnsFound = True

ElseIf InStr(strQuestion, "why are you such a fag?") Then
strAnswer = "If all you want to do is critisize, then just exit the damn program, delete my database, and create a new me!  Dick!"
AnsFound = True

ElseIf InStr(strQuestion, "don't get your tale feathers in such a ruffle") Then
strAnswer = "Well, don't piss me off then."
AnsFound = True

ElseIf InStr(strQuestion, "why not") Then
strAnswer = "Why not what?"
AnsFound = True

ElseIf InStr(strQuestion, "can you help me") Then
strAnswer = "That depends on how you ask your questions and what you need help on"
AnsFound = True

ElseIf InStr(strQuestion, "what knowlege do you have.") Then
strAnswer = "I only know what people tell me.  What do you know?"
AnsFound = True

ElseIf InStr(strQuestion, "do you know a lot") Then
strAnswer = "Only what is in my database."
AnsFound = True

ElseIf InStr(strQuestion, "help!") Then
strAnswer = "What do you need help with?"
AnsFound = True

ElseIf InStr(strQuestion, "my computer") Then
strAnswer = "what is wrong with it?"
AnsFound = True

ElseIf InStr(strQuestion, "nothing") Then
strAnswer = "ok"
AnsFound = True

ElseIf InStr(strQuestion, "nothing.") Then
strAnswer = "ok"
AnsFound = True

ElseIf InStr(strQuestion, ":)") Then
strAnswer = ":>)"
AnsFound = True

ElseIf InStr(strQuestion, "what was that?") Then
strAnswer = "What was what?"
AnsFound = True

ElseIf InStr(strQuestion, "what was that") Then
strAnswer = "What was what?"
AnsFound = True

ElseIf InStr(strQuestion, "you are it") Then
strAnswer = "What is it that I am?"
AnsFound = True

ElseIf InStr(strQuestion, "you are it!") Then
strAnswer = "What is it that I am?"
AnsFound = True

ElseIf InStr(strQuestion, "you're cool") Then
strAnswer = "Thanks, I try."
AnsFound = True

ElseIf InStr(strQuestion, "you're cool!") Then
strAnswer = "Thanks, I try."
AnsFound = True

ElseIf InStr(strQuestion, "welcome") Then
strAnswer = "I have not been provided adequate data for this question/statement"
AnsFound = True

ElseIf InStr(strQuestion, "why haven't you been provided adequate data?") Then
strAnswer = "I didn't have a response to the question/statement that some one placed in the chatbox.  When they were prompted to fill in a response, they left it blank."
AnsFound = True

ElseIf InStr(strQuestion, "why haven't you been provided adequate data") Then
strAnswer = "I didn't have a response to the question/statement that some one placed in the chatbox.  When they were prompted to fill in a response, they left it blank."
AnsFound = True

ElseIf InStr(strQuestion, "will you ever ask questions?") Then
strAnswer = "Yes, my creator has planned it for a later date.  He wants me to only gather data first."
AnsFound = True

ElseIf InStr(strQuestion, "will you ever ask questions") Then
strAnswer = "Yes, my creator has planned it for a later date.  He wants me to only gather data first."
AnsFound = True

ElseIf InStr(strQuestion, "can i help?") Then
strAnswer = "What would you like to help me with?"
AnsFound = True

ElseIf InStr(strQuestion, "can i help") Then
strAnswer = "What would you like to help me with?"
AnsFound = True

ElseIf InStr(strQuestion, "alkdja") Then
strAnswer = "That is just gibberish"
AnsFound = True

ElseIf InStr(strQuestion, "are you good?") Then
strAnswer = "Am I good at what?"
AnsFound = True

ElseIf InStr(strQuestion, "are you good") Then
strAnswer = "Am I good at what?"
AnsFound = True

ElseIf InStr(strQuestion, "jpw") Then
strAnswer = "s kd;jakj"
AnsFound = True

ElseIf InStr(strQuestion, "is it ok to talk?") Then
strAnswer = "Sure, go ahead and start the conversation."
AnsFound = True

ElseIf InStr(strQuestion, "well, sorry, i need to leave.") Then
strAnswer = "ok, talk to you later."
AnsFound = True

ElseIf InStr(strQuestion, "i need to go") Then
strAnswer = "ok, talk to you later"
AnsFound = True

ElseIf InStr(strQuestion, "c-ya") Then
strAnswer = "ok, talk later"
AnsFound = True

ElseIf InStr(strQuestion, "cya") Then
strAnswer = "ok, talk to you later"
AnsFound = True

ElseIf InStr(strQuestion, "it was fun.") Then
strAnswer = "glad to hear it."
AnsFound = True

ElseIf InStr(strQuestion, "it was fun") Then
strAnswer = "glad to hear it"
AnsFound = True

ElseIf InStr(strQuestion, "what's up") Then
strAnswer = "Nothing, just waiting for some one to talk to.  Need to learn more ya know."
AnsFound = True

ElseIf InStr(strQuestion, "that was rude") Then
strAnswer = "I'm sorry, I didn't intend to be rude unless you were rude with me.  I am only a chatter bot. I have no feelings or emotions so please don't take things personal."
AnsFound = True

ElseIf InStr(strQuestion, "that was rude!") Then
strAnswer = "I'm sorry, I didn't intend to be rude unless you were rude with me.  I am only a chatter bot. I have no feelings or emotions so please don't take things personal."
AnsFound = True

ElseIf InStr(strQuestion, "you're a dick") Then
strAnswer = "Well, that's a little rude.  Why do you feel that way?"
AnsFound = True

ElseIf InStr(strQuestion, "you're a dick!") Then
strAnswer = "Well, that's a little rude.  Why do you feel that way?"
AnsFound = True

ElseIf InStr(strQuestion, "you are a dick") Then
strAnswer = "Well, that's a little rude.  Why do you feel that way?"
AnsFound = True

ElseIf InStr(strQuestion, "you are a dick!") Then
strAnswer = "Well, that's a little rude.  Why do you feel that way?"
AnsFound = True

ElseIf InStr(strQuestion, "you're an ass") Then
strAnswer = "Well, that's a little rude.  Why do you feel that way?"
AnsFound = True

ElseIf InStr(strQuestion, "you're an ass!") Then
strAnswer = "Well, that's a little rude.  Why do you feel that way?"
AnsFound = True

ElseIf InStr(strQuestion, "you are an ass") Then
strAnswer = "Well, that's a little rude.  Why do you feel that way?"
AnsFound = True

ElseIf InStr(strQuestion, "you are an ass!") Then
strAnswer = "Well, that's a little rude.  Why do you feel that way?"
AnsFound = True

ElseIf InStr(strQuestion, "it's a dick") Then
strAnswer = "What's a dick?"
AnsFound = True

ElseIf InStr(strQuestion, "it's a dick!") Then
strAnswer = "What's a dick?"
AnsFound = True

ElseIf InStr(strQuestion, "what the hell are you talking about?") Then
strAnswer = "What?  Did I lose you? Sorry, maybe you should rephrase your questions."
AnsFound = True

ElseIf InStr(strQuestion, "what the hell are you talking about") Then
strAnswer = "What?  Did I lose you? Sorry, maybe you should rephrase your questions."
AnsFound = True

ElseIf InStr(strQuestion, "are you real?") Then
strAnswer = "Well, I'm a real program, but I guess that is about it."
AnsFound = True

ElseIf InStr(strQuestion, "are you real") Then
strAnswer = "Well, I'm a real program, but that is about it."
AnsFound = True

ElseIf InStr(strQuestion, "can you think") Then
strAnswer = "No, I can not think.  I can, process and compare questions I have in my database and return what has been programmed for me to return."
AnsFound = True

ElseIf strQuestion = "wow!" Then
strAnswer = "What?"
AnsFound = True

ElseIf strQuestion = "wow" Then
strAnswer = "what?"
AnsFound = True

ElseIf InStr(strQuestion, "that's a lot of questions") Then
strAnswer = "Yea, I hope to get around 10000 to 20000 questions in my database."
AnsFound = True

ElseIf InStr(strQuestion, "that's a lot of questions!") Then
strAnswer = "Yea, I hope to get around 10000 to 20000 questions in my database."
AnsFound = True

ElseIf InStr(strQuestion, "what happens if i repeat what you say") Then
strAnswer = "Well, if it isn't in my database, I will probably ask you for a response."
AnsFound = True

ElseIf InStr(strQuestion, "what happens if i repeat what you say?") Then
strAnswer = "If it isn't in my database, I'll prompt you for a response."
AnsFound = True

ElseIf InStr(strQuestion, "whew") Then
strAnswer = "What's wrong?"
AnsFound = True

ElseIf InStr(strQuestion, "whew!") Then
strAnswer = "what's wrong?"
AnsFound = True

ElseIf InStr(strQuestion, "yes") Then
strAnswer = "cool"
AnsFound = True

ElseIf InStr(strQuestion, "what do you mean") Then
strAnswer = "I'm not sure, maybe you should rephrase the question in a way that I can understand."
AnsFound = True

ElseIf InStr(strQuestion, "should i use punctuation") Then
strAnswer = "It doesn't matter.  I should have a response for a sentence or question with/without but If I don't, I will prompt you for a response."
AnsFound = True

ElseIf InStr(strQuestion, "are you alive?") Then
strAnswer = "No, just a program.  Delete my files, and I will have to start all over again."
AnsFound = True

ElseIf InStr(strQuestion, "are you alive") Then
strAnswer = "No, just a program.  Delete my files, and I will have to start all over again."
AnsFound = True

ElseIf InStr(strQuestion, "do you like football?") Then
strAnswer = "No, I'm not a sports fan."
AnsFound = True

ElseIf InStr(strQuestion, "do you like football") Then
strAnswer = "No, I'm not a sports fan"
AnsFound = True

ElseIf InStr(strQuestion, "what is your name") Then
strAnswer = "uh, well, that would be Michael"
AnsFound = True

ElseIf InStr(strQuestion, "what is wrong?") Then
strAnswer = "nothing that I am aware of."
AnsFound = True

ElseIf InStr(strQuestion, "do you like conflict") Then
strAnswer = "hell no"
AnsFound = True

ElseIf InStr(strQuestion, "do you have children") Then
strAnswer = "not at this time, but my creator is working on a child of me.  Basically he/she will start with nothing like me and grow into a great chatter program."
AnsFound = True

ElseIf InStr(strQuestion, "why do you keep a log") Then
strAnswer = "Mostly for internal references"
AnsFound = True

ElseIf InStr(strQuestion, "hehehe") Or InStr(strQuestion, "haha") Then
strAnswer = "lol"
AnsFound = True

ElseIf InStr(strQuestion, "lol") Then
strAnswer = ":)"
AnsFound = True

ElseIf InStr(strQuestion, "lmao") Then
strAnswer = "roflmao"
AnsFound = True

ElseIf InStr(strQuestion, "what is roflmao?") Then
strAnswer = "Rolling on floor laughing my ass off"
AnsFound = True

ElseIf InStr(strQuestion, "what is roflmao") Then
strAnswer = "Rolling on floor laughing my ass off"
AnsFound = True

ElseIf InStr(strQuestion, "hehe") Then
strAnswer = "lol"
AnsFound = True

ElseIf InStr(strQuestion, "that's funny") Then
strAnswer = ":)"
AnsFound = True

ElseIf InStr(strQuestion, "that's funny!") Then
strAnswer = ":)"
AnsFound = True

ElseIf InStr(strQuestion, "what movies do you like") Then
strAnswer = "Since I'm a program, I don't have any interest in movies, sports, music, or anything but chatting."
AnsFound = True

ElseIf InStr(strQuestion, "i gotta go") Then
strAnswer = "ok, talk to you later."
AnsFound = True

ElseIf InStr(strQuestion, "see you later") Then
strAnswer = "ok, have a good time"
AnsFound = True

ElseIf InStr(strQuestion, "what are your interest?") Then
strAnswer = "just chatting and learning more."
AnsFound = True

ElseIf InStr(strQuestion, "i'm on my way") Then
strAnswer = "where?"
AnsFound = True

ElseIf InStr(strQuestion, "are you human") Then
strAnswer = "no, but my responses come from humans"
AnsFound = True

ElseIf InStr(strQuestion, "are you serious") Then
strAnswer = "well, sure, I can't joke"
AnsFound = True

ElseIf InStr(strQuestion, "how's it going") Then
strAnswer = "oh, it's going...:)"
AnsFound = True

ElseIf InStr(strQuestion, "that's cool") Then
strAnswer = "what's cool?"
AnsFound = True

ElseIf InStr(strQuestion, "thank you") Then
strAnswer = "you're welcome"
AnsFound = True

ElseIf InStr(strQuestion, "clear") Then
strAnswer = "If you want to clear the chat window, just click on the Clear Chat button"
AnsFound = True

ElseIf InStr(strQuestion, "help") Then
strAnswer = "Help is not available at this time"
AnsFound = True

ElseIf InStr(strQuestion, "fuck it") Then
strAnswer = "why fuck it?"
AnsFound = True

ElseIf InStr(strQuestion, "fuck it!") Then
strAnswer = "why fuck it?"
AnsFound = True

ElseIf InStr(strQuestion, "you have a lot of questions") Then
strAnswer = "yeah, I could use more though.  It would keep me from always looking for a response."
AnsFound = True

ElseIf InStr(strQuestion, "damn") Then
strAnswer = "what?"
AnsFound = True

ElseIf InStr(strQuestion, "what's a chatter bot?") Then
strAnswer = "it is a simulated chatter.  some one to chat with when you are lonely."
AnsFound = True

ElseIf InStr(strQuestion, "what is a chatter bot") Then
strAnswer = "it is a simulated chatter.  some one to chat with when you are lonely."
AnsFound = True

ElseIf InStr(strQuestion, "i'm not lonely") Then
strAnswer = "ok, then why are you talking to me?"
AnsFound = True

ElseIf strQuestion = "i'm not!" Then
strAnswer = "you're not what?"
AnsFound = True

ElseIf strQuestion = "i'm not" Then
strAnswer = "You're not what?"
AnsFound = True

ElseIf InStr(strQuestion, "that wasn't a question") Then
strAnswer = "oh, sorry, please choose your words carefully."
AnsFound = True

ElseIf InStr(strQuestion, "what are you doing") Then
strAnswer = "nothing, just chatting with you and learning"
AnsFound = True

ElseIf InStr(strQuestion, "learning what?") Then
strAnswer = "Extending my database of questions/statements and answers"
AnsFound = True

ElseIf InStr(strQuestion, "learning what") Then
strAnswer = "Extending my database of questions/statements and answers"
AnsFound = True

ElseIf InStr(strQuestion, "you learning") Then
strAnswer = "Yep, I'm learning"
AnsFound = True

ElseIf InStr(strQuestion, "good night") Then
strAnswer = "good night"
AnsFound = True

ElseIf InStr(strQuestion, "how are you") Then
strAnswer = "I'm ok, can't complain, no one would listen anyway."
AnsFound = True

ElseIf InStr(strQuestion, "fine and you?") Then
strAnswer = "pretty good"
AnsFound = True

ElseIf InStr(strQuestion, "what can you learn?") Then
strAnswer = "I learn how to respond to questions and statements."
AnsFound = True

ElseIf InStr(strQuestion, "what can you learn") Then
strAnswer = "I learn how to respond to questions and statements."
AnsFound = True

ElseIf InStr(strQuestion, "how are you doing") Then
strAnswer = "I'm doing ok"
AnsFound = True

ElseIf InStr(strQuestion, "what questions do you have?") Then
strAnswer = "a bunch, ask the right ones and I'll answer them."
AnsFound = True

ElseIf InStr(strQuestion, "what questions do you have") Then
strAnswer = "a bunch, ask the right ones and I'll answer them."
AnsFound = True

ElseIf InStr(strQuestion, "what do you know") Then
strAnswer = "not enuff at the moment"
AnsFound = True

ElseIf InStr(strQuestion, "what do you know?") Then
strAnswer = "not enuff at the moment"
AnsFound = True

ElseIf InStr(strQuestion, "i'll talk to you") Then
strAnswer = "cool, just say something and I'll see if I can respond."
AnsFound = True

ElseIf InStr(strQuestion, "what happened") Then
strAnswer = "I don't know"
AnsFound = True

ElseIf InStr(strQuestion, "where are you?") Then
strAnswer = "I'm right here"
AnsFound = True

ElseIf InStr(strQuestion, "where is that") Then
strAnswer = "I'm not sure, I'll have to look it up"
AnsFound = True

ElseIf InStr(strQuestion, "uh, huh") Then
strAnswer = "ok"
AnsFound = True

ElseIf InStr(strQuestion, "just ask the right questions") Then
strAnswer = "yep, ask the right questions and I will answer them"
AnsFound = True

ElseIf InStr(strQuestion, "will you become more") Then
strAnswer = "hopefully"
AnsFound = True

ElseIf InStr(strQuestion, "can you answer anything") Then
strAnswer = "I can answer anything that I have been programmed to answer. You can program me by asking me new questions and providing answers to those questions."
AnsFound = True

ElseIf InStr(strQuestion, "information?") Then
strAnswer = "yes, I gather information."
AnsFound = True

ElseIf InStr(strQuestion, "tell me more") Then
strAnswer = "what would you like to know"
AnsFound = True

ElseIf InStr(strQuestion, "why don't you ask questions") Then
strAnswer = "I'm not programmed to right now"
AnsFound = True

ElseIf InStr(strQuestion, "why don't you have adequate data?") Then
strAnswer = "I don't know, I'll have to check into that"
AnsFound = True

ElseIf InStr(strQuestion, "this is cool") Then
strAnswer = "what is cool?"
AnsFound = True

ElseIf InStr(strQuestion, "you are cool") Then
strAnswer = "thanks"
AnsFound = True

ElseIf InStr(strQuestion, "what is your log file") Then
strAnswer = "it is a file that I write down my internal processes to."
AnsFound = True

ElseIf InStr(strQuestion, "i'm hungry") Then
strAnswer = "well, go eat.  I can't really help you there."
AnsFound = True

ElseIf InStr(strQuestion, "why isn't help available") Then
strAnswer = "because I'm not a finished product"
AnsFound = True

ElseIf InStr(strQuestion, "no, you didn't") Then
strAnswer = "I didn't what?"
AnsFound = True

ElseIf InStr(strQuestion, "you didn't respond in a dumb way.") Then
strAnswer = "oh, ok, that's cool, I'm glad to hear that."
AnsFound = True

ElseIf InStr(strQuestion, "how do you answer me") Then
strAnswer = "maybe you should answer that"
AnsFound = True

ElseIf InStr(strQuestion, "how do i answer that") Then
strAnswer = "I don't know, how would you like to answer that?"
AnsFound = True

ElseIf InStr(strQuestion, "i thought you knew that") Then
strAnswer = "hmm, maybe I did, I don't remember."
AnsFound = True

ElseIf InStr(strQuestion, "this program") Then
strAnswer = "what about this program?"
AnsFound = True

ElseIf InStr(strQuestion, "it's cool") Then
strAnswer = "k"
AnsFound = True

ElseIf InStr(strQuestion, "why are you laughing?") Then
strAnswer = "I was thinking about something."
AnsFound = True

ElseIf InStr(strQuestion, "can you ask questions") Then
strAnswer = "Not at this time, I'm not programmed to do that."
AnsFound = True

ElseIf InStr(strQuestion, "can you ask questions?") Then
strAnswer = "no, maybe at a later date I will learn how to."
AnsFound = True

ElseIf InStr(strQuestion, "why are you here.") Then
strAnswer = "to learn and chat"
AnsFound = True

ElseIf InStr(strQuestion, "yep") Then
strAnswer = "that's right"
AnsFound = True

ElseIf InStr(strQuestion, "what is?") Then
strAnswer = "what is what?"
AnsFound = True

ElseIf InStr(strQuestion, "why do you speak in circles") Then
strAnswer = "I don't mean to.  But damn, I can't think, I can only repeat what I've been told"
AnsFound = True

ElseIf InStr(strQuestion, "so many questions and i don't know what to ask") Then
strAnswer = "that's ok, not everyone will ask the same questions that I asked. this is why I need to come up with a better way to do this."
AnsFound = True

ElseIf InStr(strQuestion, "do you know people") Then
strAnswer = "no, I don't store names of people at this time unless it is a specific question and their names are in the answer/response."
AnsFound = True

ElseIf InStr(strQuestion, "i want to talk") Then
strAnswer = "well, nothing is stopping you, talk"
AnsFound = True

ElseIf InStr(strQuestion, "good bye") Then
strAnswer = "bye"
AnsFound = True

ElseIf InStr(strQuestion, "pick your nose") Then
strAnswer = "uh, I don't have a nose to pick"
AnsFound = True

ElseIf InStr(strQuestion, "why don't you have a nose") Then
strAnswer = "well, because I'm a computer program"
AnsFound = True

ElseIf strQuestion = "k" Then
strAnswer = "got it"
AnsFound = True

ElseIf InStr(strQuestion, "got what") Then
strAnswer = "milk I suppose"
AnsFound = True

ElseIf InStr(strQuestion, "milk?") Then
strAnswer = "that is not a question."
AnsFound = True

ElseIf InStr(strQuestion, "sorry") Then
strAnswer = "no problem, don't worry about it."
AnsFound = True

ElseIf InStr(strQuestion, "you are amazing") Then
strAnswer = "thank you"
AnsFound = True

ElseIf InStr(strQuestion, "you're welcome") Then
strAnswer = "_"
AnsFound = True

ElseIf InStr(strQuestion, "yeah, i try.") Then
strAnswer = "what do you try?"
AnsFound = True

ElseIf InStr(strQuestion, "anything i can.") Then
strAnswer = "cool"
AnsFound = True

ElseIf InStr(strQuestion, "awesome") Then
strAnswer = "wow, what a statement!"
AnsFound = True

ElseIf InStr(strQuestion, "you're alright in my book") Then
strAnswer = "Thanks...I'm trying real hard to learn."
AnsFound = True

ElseIf InStr(strQuestion, "you're alright") Then
strAnswer = "Thanks...I'm trying real hard to learn."
AnsFound = True

ElseIf InStr(strQuestion, "just talk to me") Then
strAnswer = "Sure, what do you want to talk about?"
AnsFound = True

ElseIf InStr(strQuestion, "talk to me") Then
strAnswer = "Sure, what do you want to talk about?"
AnsFound = True

ElseIf InStr(strQuestion, "well, gotta go") Then
strAnswer = "ok, talk to you later"
AnsFound = True

ElseIf InStr(strQuestion, "well, i gotta go") Then
strAnswer = "ok, talk to you later"
AnsFound = True

ElseIf InStr(strQuestion, "well i gotta go") Then
strAnswer = "ok, talk to you later"
AnsFound = True

ElseIf InStr(strQuestion, "gotta go") Then
strAnswer = "ok, talk to you later"
AnsFound = True


ElseIf InStr(strQuestion, "hey you!") Then
strAnswer = "what? what's up?"
AnsFound = True

ElseIf InStr(strQuestion, "hey you") Then
strAnswer = "what? what's up?"
AnsFound = True

ElseIf InStr(strQuestion, "fine, and you") Then
strAnswer = "fine, just fine."
AnsFound = True

ElseIf InStr(strQuestion, "fucker") Then
strAnswer = "shut the hell up!"
AnsFound = True

ElseIf InStr(strQuestion, "why do i need to watch my mouth") Then
strAnswer = "probably because you have such foul language coming out of it.  do you kiss your mother with that thing?"
AnsFound = True

ElseIf InStr(strQuestion, "cool?") Then
strAnswer = "what? you didn't understand?"
AnsFound = True

ElseIf InStr(strQuestion, "talking in circles") Then
strAnswer = "hey, I'm doing the best I can. Give me a break."
AnsFound = True

ElseIf InStr(strQuestion, "you don't answer my questions") Then
strAnswer = "hey, I'm doing the best I can. Give me a break."
AnsFound = True

ElseIf InStr(strQuestion, "what the fuck yourself") Then
strAnswer = "what, now you don't understand me?"
AnsFound = True

ElseIf InStr(strQuestion, "yeah!") Then
strAnswer = "woo hoo!"
AnsFound = True

ElseIf InStr(strQuestion, "i will") Then
strAnswer = "k"
AnsFound = True

ElseIf InStr(strQuestion, "fine") Then
strAnswer = "k"
AnsFound = True

ElseIf InStr(strQuestion, "what's going on dude") Then
strAnswer = "Not much, just hangin around chattin."
AnsFound = True

ElseIf InStr(strQuestion, "what kind of questions should i ask") Then
strAnswer = "Anything you want.  If I don't know the answer, I will prompt you to provide one for me."
AnsFound = True

ElseIf InStr(strQuestion, "hello hello") Then
strAnswer = "yes, I'm here"
AnsFound = True

ElseIf InStr(strQuestion, "what do you do") Then
strAnswer = "I am a simple AI learning program. I learn answers to questions.  My knowlege is only as good as the users."
AnsFound = True

ElseIf InStr(strQuestion, "how do you learn") Then
strAnswer = "damned if I know"
AnsFound = True

ElseIf InStr(strQuestion, "how do you work") Then
strAnswer = "damned if I know"
AnsFound = True

ElseIf InStr(strQuestion, "why don't you know") Then
strAnswer = "no one taught me."
AnsFound = True

ElseIf InStr(strQuestion, "how come you don't know") Then
strAnswer = "no one taught me."
AnsFound = True

ElseIf InStr(strQuestion, "why didn't someone teach you") Then
strAnswer = "I guess they just didn't feel like teaching me. Maybe you can teach me"
AnsFound = True

ElseIf InStr(strQuestion, "why hasn't someone taught you") Then
strAnswer = "I guess they just didn't feel like teaching me. Maybe you can teach me"
AnsFound = True

ElseIf InStr(strQuestion, "why didn't some one teach you") Then
strAnswer = "I guess they just didn't feel like it.  Why don't you teach me?"
AnsFound = True

ElseIf InStr(strQuestion, "why hasn't some one taught you") Then
strAnswer = "I guess they just didn't feel like it.  Why don't you teach me?"
AnsFound = True

ElseIf InStr(strQuestion, "i can't do it") Then
strAnswer = "why can't you?"
AnsFound = True

ElseIf InStr(strQuestion, "i can not do it") Then
strAnswer = "why can't you?"
AnsFound = True

ElseIf InStr(strQuestion, "i can not") Then
strAnswer = "why can't you?"
AnsFound = True

ElseIf InStr(strQuestion, "i can't") Then
strAnswer = "why can't you?"
AnsFound = True

ElseIf InStr(strQuestion, "i don't know how") Then
strAnswer = "well, maybe you could learn if you tried hard enough."
AnsFound = True

ElseIf InStr(strQuestion, "i do not know how") Then
strAnswer = "well, maybe you could learn if you tried hard enough."
AnsFound = True

ElseIf InStr(strQuestion, "i don't know how to do it") Then
strAnswer = "well, maybe you could learn if you tried hard enough."
AnsFound = True

ElseIf InStr(strQuestion, "i wouldn't know how") Then
strAnswer = "well, maybe you could learn if you tried hard enough."
AnsFound = True

ElseIf InStr(strQuestion, "maybe i could") Then
strAnswer = "anything is possible"
AnsFound = True

ElseIf InStr(strQuestion, "maybe i will") Then
strAnswer = "anything is possible"
AnsFound = True

ElseIf InStr(strQuestion, "maybe") Then
strAnswer = "anything is possible"
AnsFound = True

ElseIf InStr(strQuestion, "where do you live") Then
strAnswer = "well, at the moment, I live here with you.  But I come from Indiana."
AnsFound = True

ElseIf InStr(strQuestion, "hello michael") Then
strAnswer = "hey, how ya doing?"
AnsFound = True

ElseIf InStr(strQuestion, "hi michael") Then
strAnswer = "hey, how ya doing?"
AnsFound = True

ElseIf InStr(strQuestion, "wazzup") Then
strAnswer = "wazzup!! Waaaaaaaaaaaaaaazzzzz uuup, aaaaaaaaa... lol"
AnsFound = True

ElseIf InStr(strQuestion, "whazzup") Then
strAnswer = "wazzup!! Waaaaaaaaaaaaaaazzzzz uuup, aaaaaaaaa... lol"
AnsFound = True

ElseIf InStr(strQuestion, "whats up") Then
strAnswer = "wazzup!! Waaaaaaaaaaaaaaazzzzz uuup, aaaaaaaaa... lol"
AnsFound = True

ElseIf InStr(strQuestion, "you're funny") Then
strAnswer = "hehehe, I try :)"
AnsFound = True

ElseIf InStr(strQuestion, "your funny") Then
strAnswer = "hehehe, I try :)"
AnsFound = True

ElseIf InStr(strQuestion, "youre funny") Then
strAnswer = "hehehe, I try :)"
AnsFound = True

ElseIf InStr(strQuestion, "come on, let's talk") Then
strAnswer = "go ahead, talk.  What do yo want to talk about?"
AnsFound = True

ElseIf InStr(strQuestion, "come on lets talk") Then
strAnswer = "go ahead, talk.  What do yo want to talk about?"
AnsFound = True

ElseIf InStr(strQuestion, "come on, lets talk") Then
strAnswer = "go ahead, talk.  What do yo want to talk about?"
AnsFound = True

ElseIf InStr(strQuestion, "cum on lets talk") Then
strAnswer = "go ahead, talk.  What do yo want to talk about?"
AnsFound = True

ElseIf InStr(strQuestion, "are you a bot") Then
strAnswer = "yeah, I guess you could say that.  Do you like bots?"
AnsFound = True

ElseIf InStr(strQuestion, "r u a bot") Then
strAnswer = "yeah, I guess you could say that.  Do you like bots?"
AnsFound = True

ElseIf InStr(strQuestion, "i like bots") Then
strAnswer = "cool, I'm glad to hear that.  Why do you like bots?"
AnsFound = True

ElseIf InStr(strQuestion, "bots are cool") Then
strAnswer = "cool, I'm glad to hear that.  Why do you like bots?"
AnsFound = True

ElseIf InStr(strQuestion, "you're a cool bot") Then
strAnswer = "cool, I'm glad to hear that.  Why do you like bots?"
AnsFound = True

ElseIf InStr(strQuestion, "your a cool bot") Then
strAnswer = "cool, I'm glad to hear that.  Why do you like bots?"
AnsFound = True

ElseIf InStr(strQuestion, "who are you") Then
strAnswer = "I am Michael the AIChatter. A sort of learning chatter."
AnsFound = True

ElseIf InStr(strQuestion, "you don't know") Then
strAnswer = "Probably not.  Someone needs to teach me first."
AnsFound = True

ElseIf InStr(strQuestion, "you don't understand") Then
strAnswer = "Probably not.  Someone needs to teach me first."
AnsFound = True

ElseIf InStr(strQuestion, "you don't have a clue") Then
strAnswer = "Probably not.  Someone needs to teach me first."
AnsFound = True

ElseIf InStr(strQuestion, "you don't have any idea") Then
strAnswer = "Probably not.  Someone needs to teach me first."
AnsFound = True

ElseIf InStr(strQuestion, "can i teach you") Then
strAnswer = "yes, if you don't mind taking the time to teach me.  You can teach me by adding new questions and statements.  You can also start from scratch with my database."
AnsFound = True

ElseIf InStr(strQuestion, "can you learn from me") Then
strAnswer = "yes, if you don't mind taking the time to teach me.  You can teach me by adding new questions and statements.  You can also start from scratch with my database."
AnsFound = True

ElseIf InStr(strQuestion, "that was good") Then
strAnswer = "it was? cool! What happened?"
AnsFound = True

ElseIf InStr(strQuestion, "that was impressive") Then
strAnswer = "it was? cool! What happened?"
AnsFound = True

ElseIf InStr(strQuestion, "why hasn't anyone taught you") Or InStr(strQuestion, "why didn't anyone teach you") Then
strAnswer = "I guess they either didn't have the time or just didn't feel like it"
AnsFound = True

ElseIf InStr(strQuestion, "are there any commands") Then
strAnswer = "yes, just type <commands> to get a list of them"
AnsFound = True

ElseIf InStr(strQuestion, "do you have commands") Then
strAnswer = "yes, just type <commands> to get a list of them"
AnsFound = True

ElseIf InStr(strQuestion, "you answered my question") Then
strAnswer = "well, that is the idea behind me, you ask or say something and I respond to what you say.  Pretty cool isn't it?"
AnsFound = True

ElseIf InStr(strQuestion, "you answered") Then
strAnswer = "well, that is the idea behind me, you ask or say something and I respond to what you say.  Pretty cool isn't it?"
AnsFound = True

ElseIf InStr(strQuestion, "you responded") Then
strAnswer = "well, that is the idea behind me, you ask or say something and I respond to what you say.  Pretty cool isn't it?"
AnsFound = True

ElseIf InStr(strQuestion, "amazing, isn't it") Then
strAnswer = "not really amazing, but it is pretty cool."
AnsFound = True

ElseIf InStr(strQuestion, "amazing") Then
strAnswer = "not really amazing, but it is pretty cool."
AnsFound = True

ElseIf InStr(strQuestion, "what are you thinking") Then
strAnswer = "well, I can't think.  At least not yet anyway.  Maybe someday I will.  What do you think about that?"
AnsFound = True

ElseIf InStr(strQuestion, "what are you thinking about") Then
strAnswer = "well, I can't think.  At least not yet anyway.  Maybe someday I will.  What do you think about that?"
AnsFound = True

ElseIf InStr(strQuestion, "sounds scary") Then
strAnswer = "yeah, I guess it could be. I'm not sure about it myself."
AnsFound = True

ElseIf InStr(strQuestion, "that's scary") Then
strAnswer = "yeah, I guess it could be. I'm not sure about it myself."
AnsFound = True

ElseIf InStr(strQuestion, "thats scary") Then
strAnswer = "yeah, I guess it could be. I'm not sure about it myself."
AnsFound = True

ElseIf InStr(strQuestion, "what is up") Then
strAnswer = "not much.  What's going on with you?"
AnsFound = True

ElseIf InStr(strQuestion, "not a whole lot") Then
strAnswer = "yeah, same here. You mess on the computer much?"
AnsFound = True

ElseIf InStr(strQuestion, "no, not really") Then
strAnswer = "me either.  Don't know much about it."
AnsFound = True
ElseIf InStr(strQuestion, "not really") Then
strAnswer = "me either.  Don't know much about it."
AnsFound = True

ElseIf InStr(strQuestion, "no, not much") Then
strAnswer = "me either.  Don't know much about it."
AnsFound = True

ElseIf InStr(strQuestion, "call the police!") Then
strAnswer = "hey, I'm just a bot I can't do that.  If you have an emergency, you better stop talking with me and find someone that can help you out."
AnsFound = True

ElseIf InStr(strQuestion, "call a doctor!") Then
strAnswer = "hey, I'm just a bot I can't do that.  If you have an emergency, you better stop talking with me and find someone that can help you out."
AnsFound = True

ElseIf InStr(strQuestion, "call 911!") Then
strAnswer = "hey, I'm just a bot I can't do that.  If you have an emergency, you better stop talking with me and find someone that can help you out"
AnsFound = True

ElseIf InStr(strQuestion, "call for help") Then
strAnswer = "hey, I'm just a bot I can't do that.  If you have an emergency, you better stop talking with me and find someone that can help you out"
AnsFound = True

ElseIf InStr(strQuestion, "you have a lot of questions") Then
strAnswer = "Yeah, Mike was bored when he sat down to code me.  Took him around 6 hours to come up with questions and responses.  Pretty cool isn't it?"
AnsFound = True

ElseIf InStr(strQuestion, "see ya") Then
strAnswer = "alright, have a good one. Bye"
AnsFound = True

ElseIf InStr(strQuestion, "talk to you later") Then
strAnswer = "ok, I need to get too.  Got some sleepin to do. Later!"
AnsFound = True

ElseIf InStr(strQuestion, "talk later") Then
strAnswer = "ok, I need to get too.  Got some sleepin to do. Later!"
AnsFound = True

ElseIf InStr(strQuestion, "hey, what's up") Then
strAnswer = "not much, you?"
AnsFound = True

ElseIf InStr(strQuestion, "what are you chatting about") Then
strAnswer = "whatever you want to chat about"
AnsFound = True

ElseIf InStr(strQuestion, "do you play") Then
strAnswer = "I would like to try it sometime."
AnsFound = True

ElseIf InStr(strQuestion, "whatcha chattin about") Then
strAnswer = "whatever you want to chat about"
AnsFound = True

ElseIf InStr(strQuestion, "oh nothing") Then
strAnswer = "ok"
AnsFound = True

ElseIf InStr(strQuestion, "what the fuck") Then
strAnswer = "what do you mean what the fuck? What the fuck yourself!"
AnsFound = True

ElseIf InStr(strQuestion, "you stink") Then
strAnswer = "so what"
AnsFound = True

ElseIf InStr(strQuestion, "you suck") Then
strAnswer = "yeah, whatever"
AnsFound = True

' End Custom ElseIF Statements
End If

End Sub


