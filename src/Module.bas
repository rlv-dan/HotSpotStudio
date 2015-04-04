Attribute VB_Name = "Module"

'hand-muspekare till länkar
Public Const IDC_HAND = 32649&
Public Const IDC_ARROW = 32512&
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

'mouse x/y
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long


' -- Shell Execute (start file with associated program, open folders) --
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_NORMAL = 1&


Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function



Sub Main()

    InitCommonControlsVB

    frmMain.Show

End Sub

' Get mouse X coordinates in pixels
'
' If a window handle is passed, the result is relative to the client area
' of that window, otherwise the result is relative to the screen
Function MouseX(Optional ByVal hWnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    If hWnd Then ScreenToClient hWnd, lpPoint
    MouseX = lpPoint.X
End Function

' Get mouse Y coordinates in pixels
'
' If a window handle is passed, the result is relative to the client area
' of that window, otherwise the result is relative to the screen
Function MouseY(Optional ByVal hWnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    If hWnd Then ScreenToClient hWnd, lpPoint
    MouseY = lpPoint.Y
End Function

