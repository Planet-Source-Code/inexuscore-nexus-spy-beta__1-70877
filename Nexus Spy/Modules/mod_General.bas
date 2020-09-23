Attribute VB_Name = "mod_General"
Option Explicit

'// public win32 api declarations
Public Declare Function InitCommonControls Lib "ComCtl32.dll" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal sWndTitle As String, ByVal cLen As Long) As Long
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'// public constants declaration
Public Const VK_MENU = &H12
Public Const VK_SNAPSHOT = &H2C
Public Const KEYEVENTF_KEYUP = &H2
