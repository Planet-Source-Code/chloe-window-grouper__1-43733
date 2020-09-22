Attribute VB_Name = "basMAIN"
Option Explicit
Public Sub Main()
    On Error GoTo ErrorMain
    If App.PrevInstance = False Then
        frmMAIN.Show
    End If
    Exit Sub
ErrorMain:
    MsgBox Err & ":Error in call to Main()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
