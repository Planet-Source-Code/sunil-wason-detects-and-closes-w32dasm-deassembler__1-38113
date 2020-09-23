Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function sendmessagebystring Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)

Public Const WM_CLOSE = &H10
Public Const EXCEPTION_ACCESS_VIOLATION = &HC0000005

Public Function WinDasmHwnd() As Long

'Returns handle of the application
WinDasmHwnd = FindWindow("OWL_Window", vbNullString)
  
End Function 'WinDasmHwnd() As Boolean
    

