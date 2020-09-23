<div align="center">

## Get the "Real" Hand Cursor


</div>

### Description

This small piece of code will set the cursor to the hand cursor that is shown when you hover above a hyperlink. I know you can use a .RES file to load the cursor, but the cursor can be changed in the Mouse control panel, and a .RES cursor will not reflect the changes. NOTE: This is not all my code. I found the API call on PSC (had to do quite a bit of searching), and put together a little Subroutine to make it easier to do. The original source is here: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=31572&lngWId=1
 
### More Info
 
True/False value. If Hand=True, the hyperlink cursor will be displayed, if it is False, the standard pointer will be displayed.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Nathanael B](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nathanael-b.md)
**Level**          |Intermediate
**User Rating**    |4.8 (53 globes from 11 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/nathanael-b-get-the-real-hand-cursor__1-37076/archive/master.zip)

### API Declarations

```
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Const IDC_HAND = 32649&
Public Const IDC_ARROW = 32512&
```


### Source Code

```
Public Sub SetHandCur(Hand As Boolean)
  If Hand = True Then
    SetCursor LoadCursor(0, IDC_HAND)
  Else
    SetCursor LoadCursor(0, IDC_ARROW)
  End If
End Sub
```

