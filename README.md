<div align="center">

## Screenshot with print option \(BitBlt and DC's\)


</div>

### Description

This little program will let you take a picture of the screen at any time and Blitter from the screen's DC to the Form's DC making it appear as the form background you can then save it as a .BMP

or print it out. Sorry about the messy .BAS but its all there. Please take the time to leave a comment and vote. Thank you. Enjoy!  =)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-07-18 02:19:22
**By**             |[Adam Orenstein](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/adam-orenstein.md)
**Level**          |Advanced
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD78897182000\.zip](https://github.com/Planet-Source-Code/adam-orenstein-screenshot-with-print-option-bitblt-and-dc-s__1-9851/archive/master.zip)

### API Declarations

```
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Public Declare Function GetDesktopWindow Lib "user32" () As Long
```





