message=InputBox("What do you want your computer to say?","text to voice")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message 