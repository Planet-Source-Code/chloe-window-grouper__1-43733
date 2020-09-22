Attribute VB_Name = "basREMOVECAPTION"
Option Explicit
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long

Public Declare Function MoveWindow Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long

Public Const SM_CYBORDER = 6
Public Const SM_CYCAPTION = 4
Public Const GWL_STYLE = (-16)
Public Const WS_CAPTION = &HC00000
Private Const WS_BORDER = &H800000
Sub RemoveCaptionBar(hWnd As Long, Optional lngLeft, Optional lngTop)
    Dim OldStyle As Long, NewStyle As Long
    Dim r As RECT
    Dim RetVal As Long
    Dim dx As Long, dy As Long
    Dim r1 As RECT
    OldStyle = GetWindowLong(hWnd, GWL_STYLE)
    NewStyle = OldStyle And Not WS_CAPTION
    OldStyle = SetWindowLong(hWnd, GWL_STYLE, NewStyle)
    RetVal = GetWindowRect(hWnd, r)
    dx = r.Right - r.Left + 1
    dy = r.Bottom - r.Top + 1
    If IsMissing(lngLeft) Then lngLeft = r.Left
    If IsMissing(lngTop) Then lngTop = r.Top
    RetVal = MoveWindow(hWnd, lngLeft, lngTop, lngLeft + dx, dy + lngTop - (GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYBORDER) * 4), True)
End Sub
Sub AddCaptionBar(hWnd As Long, Optional lngLeft, Optional lngTop)
    Dim OldStyle As Long, NewStyle As Long
    Dim r As RECT
    Dim RetVal As Long
    Dim dx As Long, dy As Long
    Dim r1 As RECT
    OldStyle = GetWindowLong(hWnd, GWL_STYLE)
    NewStyle = OldStyle Or WS_CAPTION Or WS_BORDER
    OldStyle = SetWindowLong(hWnd, GWL_STYLE, NewStyle)
    RetVal = GetWindowRect(hWnd, r)
    dx = r.Right - r.Left + 1
    dy = r.Bottom - r.Top + 1
    If IsMissing(lngLeft) Then lngLeft = r.Left
    If IsMissing(lngTop) Then lngTop = r.Top
    RetVal = MoveWindow(hWnd, lngLeft, lngTop, lngLeft + dx, dy + lngTop - (GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYBORDER) * 4), True)
End Sub

