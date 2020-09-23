<div align="center">

## SwitchToThisWindow


</div>

### Description

Another way of changing the focus to a particular window, using an un-documented API called SwitchToThisWindow.This API works on Windows 3x, Windows 9x/ME and Windows 2000. I have found it much more reliable than Setfocus/SetFocusAPI or SetForegroundWindow
 
### More Info
 
Window caption

I have defined a public function ( GetMessageWindow ) in order to show how the SwitchToThisWindow API works in conjunction with the FindWindow API. The GetMessageWindow function can of course be improved to accept parameters and make it more flexible. For this example I just use it to switch to a standard message box.

Long - Window handle on success, zero on failure


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Simon Morgan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/simon-morgan.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/simon-morgan-switchtothiswindow__1-43043/archive/master.zip)

### API Declarations

See code window


### Source Code

```
' Library imports
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SwitchToThisWindow Lib "user32" (ByVal hWnd As Long, ByVal hWindowState As Long) As Long
' Function to locate and focus on a MsgBox
Public Function GetMessageWindow() As Long
Dim hMessageBox As Long
' First get the message box's handle
hMessageBox& = FindWindow("#32770", vbNullString)
If hMessageBox Then
 ' Set focus on the message box
 GetMessageWindow& = SwitchToThisWindow
  (hMessageBox, vbNormalFocus)
Else
 GetMessageWindow& = 0
End If
End Function
' Calling the GetMessageWindow function
RetVal& = GetMessageWindow()
If RetVal& Then
 ' Do something like SendKeys{Enter}
End If
```

