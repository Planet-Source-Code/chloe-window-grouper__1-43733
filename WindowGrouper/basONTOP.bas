Attribute VB_Name = "basONTOP"
Option Explicit
'Constants for topmost.
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Enum ONTOPSETTING
    WINDOW_ONTOP = HWND_TOPMOST
    WINDOW_NOT_ONTOP = HWND_NOTOPMOST
End Enum
'------------------------------------------------------------
' Author:  Clint M. LaFever [clint.m.lafever@cpmx.saic.com]
' Purpose:  Functionality to Set a window always on top or turn it off.
' Date: March,10 1999 @ 10:18:37
'------------------------------------------------------------
Public Sub SetFormOnTop(formHWND As Long, Optional sSETTING As ONTOPSETTING = WINDOW_ONTOP)
    On Error Resume Next
    Call SetWindowPos(formHWND, sSETTING, 0, 0, 0, 0, FLAGS)
End Sub

