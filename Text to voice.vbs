message=InputBox("Type anything in the box and it will be said.", "Text2Voice")
      Set sapi = CreateObject("sapi.spvoice")
sapi.Speak message 
