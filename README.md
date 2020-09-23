<div align="center">

## Create a Window without a Form \!\!\!


</div>

### Description

This module demonstrates how to generate a Window using the API. Why use the Visual Basic FormDesigner when you can use the API?

Okay it's much easier to

use the Designer, but a good VB-Developer has to see and understand a module like this ;-).

I have translated the C++ - Code (MSDN\SDK) to VB.
 
### More Info
 
Create a new Project and add a module. Remove Form1. Add the following API Declarations and the code.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Henning Tillmann](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/henning-tillmann.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/henning-tillmann-create-a-window-without-a-form__1-7171/archive/master.zip)

### API Declarations

```

Private Declare Function apiRegisterClass Lib "user32" _
             Alias "RegisterClassA" _
            (Class As WNDCLASS) As Long
Private Declare Function apiCreateWindowEx Lib "user32" _
             Alias "CreateWindowExA" _
            (ByVal dwExStyle As Long, _
             ByVal lpClassName As String, _
             ByVal lpWindowName As String, _
             ByVal dwStyle As Long, _
             ByVal x As Long, _
             ByVal y As Long, _
             ByVal nWidth As Long, _
             ByVal nHeight As Long, _
             ByVal hWndParent As Long, _
             ByVal hMenu As Long, _
             ByVal hInstance As Long, _
             lpParam As Any) As Long
Private Declare Function apiLoadIcon Lib "user32" _
             Alias "LoadIconA" _
            (ByVal hInstance As Long, _
             ByVal lpIconName As String) As Long
Private Declare Function apiLoadCursor Lib "user32" _
             Alias "LoadCursorA" _
            (ByVal hInstance As Long, _
             ByVal lpCursorName As String) As Long
Private Declare Function apiDispatchMessage Lib "user32" _
             Alias "DispatchMessageA" _
            (lpMsg As MSG) As Long
Private Declare Function apiGetMessage Lib "user32" _
             Alias "GetMessageA" _
            (lpMsg As MSG, _
             ByVal hWnd As Long, _
             ByVal wMsgFilterMin As Long, _
             ByVal wMsgFilterMax As Long) As Long
Private Declare Function apiDefWindowProc Lib "user32" _
             Alias "DefWindowProcA" _
            (ByVal hWnd As Long, _
             ByVal wMsg As Long, _
             ByVal wParam As Long, _
             ByVal lParam As Long) As Long
Private Declare Function apiSetWindowPos Lib "user32" _
             Alias "SetWindowPos" _
            (ByVal hWnd As Long, _
             ByVal hWndInsertAfter As Long, _
             ByVal x As Long, _
             ByVal y As Long, _
             ByVal cx As Long, _
             ByVal cy As Long, _
             ByVal wFlags As Long) As Long
Private Declare Function apiUnregisterClass Lib "user32" _
             Alias "UnregisterClassA" _
            (ByVal lpClassName As String, _
             ByVal hInstance As Long) As Long
Private Type WNDCLASS
  style As Long
  lpfnwndproc As Long
  cbClsextra As Long
  cbWndExtra2 As Long
  hInstance As Long
  hIcon As Long
  hCursor As Long
  hbrBackground As Long
  lpszMenuName As String
  lpszClassName As String
End Type
Private Type POINTAPI
  x As Long
  y As Long
End Type
Private Type MSG
  hWnd As Long
  message As Long
  wParam As Long
  lParam As Long
  time As Long
  pt As POINTAPI
End Type
Private Const CS_OWNDC = &H20
Private Const CS_VREDRAW = &H1
Private Const CS_HREDRAW = &H2
Private Const IDI_APPLICATION = 32512&
Private Const IDC_ARROW = 32512&
Private Const COLOR_WINDOW = 5
Private Const WS_OVERLAPPED = &H0&
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_HSCROLL = &H100000
Private Const WS_VSCROLL = &H200000
Const HT_CLASSNAME = "PlanetSourceCodeTest"
Const HT_WINDOWTITLE = "HELLO WORLD!"
Dim hWnd As Long
```


### Source Code

```
Sub Main()
  Dim WC As WNDCLASS
  Dim dwRetVal As Long
  Dim msgWnd As MSG
  WC.lpszClassName = HT_CLASSNAME
  WC.lpfnwndproc = GetAddressOf(AddressOf MainWndProc)
  WC.style = CS_OWNDC Or CS_VREDRAW Or CS_HREDRAW
  WC.hInstance = App.hInstance
  WC.hIcon = apiLoadIcon(0, IDI_APPLICATION)
  WC.hCursor = apiLoadCursor(0, IDC_ARROW)
  WC.hbrBackground = COLOR_WINDOW
  WC.cbClsextra = 0
  WC.cbWndExtra2 = 0
  dwRetVal = apiRegisterClass(WC)
  Debug.Print "RegisterClass returns '" & CStr(dwRetVal) & "'."
  hWnd = apiCreateWindowEx(0, HT_CLASSNAME, HT_WINDOWTITLE, WS_OVERLAPPEDWINDOW, 0, 0, 0, 0, 0, 0, App.hInstance, 0)
  Debug.Print "CreateWindowEx returns hWnd '" & CStr(hWnd) & "'."
  dwRetVal = apiSetWindowPos(hWnd, 0, 200, 200, 300, 300, &H40)
  Debug.Print "SetWindowPos returns '" & CStr(dwRetVal) & "'."
  Do While apiGetMessage(msgWnd, hWnd, 0&, 0&) > 0
   apiDispatchMessage msgWnd ': DoEvents
  Loop
  dwRetVal = apiUnregisterClass(HT_CLASSNAME, App.hInstance)
  Debug.Print "UnregisterClass returns '" & CStr(dwRetVal) & "'."
End Sub
Private Function MainWndProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  MainWndProc = apiDefWindowProc(hWnd, wMsg, wParam, lParam)
End Function
Private Function GetAddressOf(ProcAddress As Long) As Long
  GetAddressOf = ProcAddress
End Function
```

