VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLIST 
   Caption         =   "Window List"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvLIST 
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    On Error Resume Next
    InitlvLIST
End Sub
Private Sub InitlvLIST()
    On Error GoTo ErrorInitlvLIST
    With lvLIST
        .View = lvwReport
        .FullRowSelect = True
        .ColumnHeaders.Add , , "WINDOW"
        .Sorted = True
        .SortKey = 0
        .LabelEdit = lvwManual
    End With
    Exit Sub
ErrorInitlvLIST:
    MsgBox Err & ":Error in call to InitlvLIST()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    With Me
        .lvLIST.Move 0, 0, .ScaleWidth, .ScaleHeight
    End With
End Sub
Public Sub AddWindow(wCap As String, hWnd As Long)
    On Error GoTo ErrorAddWindow
    Dim itm As ListItem
    If hWnd <> Me.hWnd And hWnd <> frmMAIN.hWnd And UCase(wCap) <> "PROGRAM MANAGER" Then
        Set itm = lvLIST.ListItems.Add(, , wCap)
        itm.Tag = hWnd
        If lvLIST.ColumnHeaders(1).Width < Me.TextWidth(wCap) + 120 Then lvLIST.ColumnHeaders(1).Width = Me.TextWidth(wCap) + 120
    End If
    Exit Sub
ErrorAddWindow:
    MsgBox Err & ":Error in call to AddWindow()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub lvLIST_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next
    If lvLIST.SortOrder = lvwAscending Then
        Me.lvLIST.SortOrder = lvwDescending
    Else
        Me.lvLIST.SortOrder = lvwAscending
    End If
End Sub
Private Sub lvLIST_DblClick()
    On Error Resume Next
    Dim itm As ListItem
    Set itm = lvLIST.SelectedItem
    If Not itm Is Nothing Then
        frmMAIN.AssignChild CLng(itm.Tag), itm.Text
        lvLIST.ListItems.Remove itm.Index
    End If
    Set itm = Nothing
End Sub
