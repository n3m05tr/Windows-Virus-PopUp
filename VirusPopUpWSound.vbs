Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")

' Plays the audio file
oPlayer.URL = "distination of the wav file"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 100
Wend

x=msgbox("your message here",0+16,"your message here")

'This makes the message pop up. All of the codes for the types of pop ups:

'32 warning query icon.
'48 warning message icon
'64 information message icon

'0 normal message box
'4906 always stays on top of desktop

'16 critical message icon
'0 OK button only
'1 OK and cancel buttons
'2 abort, retry, and cancel buttons
'3 yes, no , and cancel buttons
'4 yes and no buttons
'5 retry and cancel buttons

' Stops the audio file
oPlayer.close




