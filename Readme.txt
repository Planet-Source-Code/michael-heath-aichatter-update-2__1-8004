This is another alpha release of AIChatter.  There are three subs that have been added that sort of change things.  SUB CustomElseIF, SUB KeyWords, and Sub AIRandom.

Sub CustomElseIF:
	This sub lets the programmer add questions/statements and responses that he/she
	thinks a user would ask.

Sub KeyWords:
	This is where the programmer adds keyphrases or keywords that he wants Michael
	to lock on to.  If there aren't any matching questions or statements in the general 	data file or in the customelseif, then lock on to some keyphrases or keywords in the  	sentence and come up with a resonable response.  Just take a look at the sub and it 
	will make sense.

Sub AIRandom:
	These are random responses to questions/statements that have been repeated.

Please be aware that everything isn't completely implemented yet and the code is still in progress.  Due to this, you may still get some strange and unusual responses as well as getting the same response over and over again for the majority. It takes a lot of time trying to come up with multiple responses to questions/statements.

You'll also notice a file called sgeneral.exe.  This file is used to send the general.dat file to me.  If you do not wish to send me any learned questions/statements in the dat file, then DO NOT RUN sgeneral.exe.  I would really like to have that file, but it is totally up to you.  I can't make michael respond to questions/statements that I don't know about.


I'm open to improvements and or suggestions to this code.  Please don't send any suggestions that would request michael's coding to make him recognize almost everything.  I've already started and tested michael with that sort of coding.  Someone told me not to do more than what was needed.  After starting that, I realized they were correct.  The code for making michael analyze sentences proved to be a greater task than anticipated.  I stepped into a huge mess of nested question/statements and responses.  It is very difficult to try and put an actual neuronet to michael's processing.  Just settle for the string searching and try to make it better.  I don't recommend to anyone that you try making him understand everything in the english language.  There are a couple of examples in Sub KeyWords that showed just a small problem with analyzing a sentence.  You'll know them when you see them.  They are the ones that lock on to a keyphrase and then proceed to rip the sentence apart to find exactly what the user typed.  When you look at them, just imagine trying to do that on a larger scale and I'm sure you will change your mind about ripping sentences apart.

Enjoy the code. --- Michael Heath
Direct any of your questions or comments to Michael Heath at

			mheath@indy.net

