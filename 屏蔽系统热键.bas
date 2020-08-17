Attribute VB_Name = "ÆÁ±ÎÏµÍ³ÈÈ¼ü"
Option Explicit
Public hHook As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetKeyState Lib "USER32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowsHookEx Lib "USER32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "USER32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "USER32" (ByVal hHook As Long) As Long
Public Const HC_ACTION = 0
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const VK_TAB = &H9
Public Const VK_CONTROL = &H11
Public Const VK_ESCAPE = &H1B
Public Const WH_KEYBOARD_LL = 13
Public Const LLKHF_ALTDOWN = &H20
Public Type KBDLLHOOKSTRUCT
vkCode As Long
scanCode As Long
flags As Long
time As Long
dwExtraInfo As Long
End Type
Dim p As KBDLLHOOKSTRUCT
Public Function MyKBHook(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 Dim pbj As Boolean
 If (nCode = HC_ACTION) Then
  If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Or wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
  CopyMemory p, ByVal lParam, Len(p)
  pbj = _
   ((p.vkCode = VK_TAB) And ((p.flags And LLKHF_ALTDOWN) <> 0)) Or _
   ((p.vkCode = VK_ESCAPE) And ((p.flags And LLKHF_ALTDOWN) <> 0)) Or _
   ((p.vkCode = VK_ESCAPE) And ((GetKeyState(VK_CONTROL) And &H8000) <> 0)) Or _
   ((p.vkCode = 91) Or (p.vkCode = 92) Or (p.vkCode = 93)) Or (p.vkCode = &H73)
  End If
 End If
 If pbj Then
  MyKBHook = -1
 Else
  MyKBHook = CallNextHookEx(0, nCode, wParam, ByVal lParam)
 End If
End Function

