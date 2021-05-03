Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")

' Plays the audio file
oPlayer.URL = "distination of the wav file"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 100
Wend

x=msgbox("your message here",0+16,"your message here")

' Stops the audio file
oPlayer.close




