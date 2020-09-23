<div align="center">

## Take Screenshot From VB


</div>

### Description

This will take a screenshot of the screen from inside VB. It is not mainly for taking screenshots of apps (Alt + PrintScreen or just PrintScreen). This is mainly for stuff like spying, etc. So this will have the same effect as the PrintScreen key. You will be responsible for protecting the clipboard info. Please erase the clipboard before you run this function as then it will take the previous data in the clipboard first, and then work the second time. I tried putting Clipboard.Clear into the sub but then it will not work at all.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[excelblue \(Mark Lu\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/excelblue-mark-lu.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/excelblue-mark-lu-take-screenshot-from-vb__1-33789/archive/master.zip)

### API Declarations

```
Declare Sub keybd_event Lib "user32" Alias "keybd_event" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
```


### Source Code

```
Public Sub GetScreenShot(SetObj As Object)
  On Error Resume Next
  Dim CurrCBData As Variant, CurrCBText As String, CurrPict As String
  keybd_event vbKeySnapshot, 1, 0, 0
  SetObj.Picture = Clipboard.GetData(vbCFBitmap)
  Clipboard.Clear
End Sub
```

