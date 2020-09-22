Attribute VB_Name = "basENUMWINDOWS"
Option Explicit
Private Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetParent& Lib "user32" (ByVal hWnd As Long)
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_CLOSE = &H10
Dim sPattern As String, hFind As Long
Private Listing As Boolean
Public Function CloseWindowByhWnd(hWnd As Long) As Long
    On Error GoTo ErrorCloseWindowByhWnd
    CloseWindowByhWnd = PostMessage(hWnd, WM_CLOSE, 0&, 0&)
    Exit Function
ErrorCloseWindowByhWnd:
    MsgBox Err & ":Error in CloseWindowByhWnd.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Function
End Function
Function EnumWinProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Dim k As Long, sName As String, frm As frmLIST, f As Form, Found As Boolean
    If Listing = True Then
        Found = False
        For Each f In Forms
            If f.Name = "frmLIST" Then
                Set frm = f
                Found = True
                Exit For
            End If
        Next
        If Found = False Then
            Set frm = New frmLIST
            frm.Show , frmMAIN
        End If
    End If
    If IsWindowVisible(hWnd) And GetParent(hWnd) = 0 Then
        sName = Space$(128)
        k = GetWindowText(hWnd, sName, 128)
        If k > 0 Then
            sName = Left$(sName, k)
            If lParam = 0 Then sName = UCase(sName)
            If Listing = True Then frm.AddWindow sName, hWnd
            If sName Like sPattern Then
                hFind = hWnd
                EnumWinProc = 0
                Exit Function
            End If
        End If
    End If
    EnumWinProc = 1
End Function
Public Function FindWindowWild(sWild As String, Optional bMatchCase As Boolean = True) As Long
    On Error Resume Next
    sPattern = sWild
    hFind = 0
    If Not bMatchCase Then sPattern = UCase(sPattern)
    EnumWindows AddressOf EnumWinProc, bMatchCase
    FindWindowWild = hFind
End Function
Public Sub ListWindows()
    On Error GoTo ErrorListWindows
    Dim frm As Form
    For Each frm In Forms
        If frm.Name = "frmLIST" Then
            frm.lvLIST.ListItems.Clear
            Exit For
        End If
    Next
    Listing = True
    FindWindowWild "QWERTY", False
    Listing = False
    Exit Sub
ErrorListWindows:
    Listing = True
    MsgBox Err & ":Error in call to ListWindows()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub

