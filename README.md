<div align="center">

## Capture desktop/Screen shot, and save to file, easiest way using only one API call\!\! A must see\!


</div>

### Description

This will take a screen shot using only one API call, no bitblt and other complicated routines. This is simple and works. This code will simulate a keypress (the snap shot button), then will copy the data to clipboard (like if you had pressed the Snap shot button on your keyboard) and save it to file. Vote for me or gimme' some feedback if you think this is cool.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[$mTp ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mtp.md)
**Level**          |Intermediate
**User Rating**    |4.9 (122 globes from 25 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mtp-capture-desktop-screen-shot-and-save-to-file-easiest-way-using-only-one-api-call-a-mus__1-10994/archive/master.zip)





### Source Code

'****Declares
<Font face="verdana" size ="2"><P>Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
ByVal dwFlags As Long, ByVal dwExtraInfo As Long)</P>
<P>Public Function Capture_Desktop(ByVal Destination$) as Boolean </P>
On Error goto errl
<br>DoEvents
<br>Call keybd_event(vbKeySnapshot, 1, 0, 0) 'Get the screen and copy it to clipboard
<br>DoEvents 'let computer catch up
<br>SavePicture Clipboard.GetData(vbCFBitmap), Destination$ ' saves the clipboard data to a BMP file
<br>Capture_Desktop = True
<br>Exit Function
<br>errl:
<br>Msgbox "Error number: " & err.number & ". " & err.description
<br>Capture_Desktop = False
<br>End Function
'A lil' example
<br>Private Sub Command1_Click()
<br>Capture_Desktop "c:\windows\desktop\desktop.bmp" 'That's it

