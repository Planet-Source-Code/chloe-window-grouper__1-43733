VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMAIN 
   BackColor       =   &H80000005&
   Caption         =   "Window Grouper"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMAIN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMONITOR 
      Interval        =   100
      Left            =   1440
      Top             =   2880
   End
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   3120
      Top             =   2640
   End
   Begin MSComctlLib.StatusBar sbSTATUS 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.Image imgCLOSE 
      Height          =   210
      Left            =   3960
      Picture         =   "frmMAIN.frx":030A
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFILE 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuGET 
         Caption         =   "&Get windows"
      End
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Const SW_NORMAL = 1
Private Const SW_HIDE = 0
Private Const SW_MAXIMIZE = 3

Public MonitorFor As String
Private Sub Form_Load()
    On Error Resume Next
    Me.MonitorFor = "*- microsoft internet explorer"
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = vbRightButton Then
        Me.PopupMenu mnuFILE
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Dim X As Long, hwnd As Long
    For X = 1 To Me.sbSTATUS.Panels.Count
        hwnd = CLng(Left(sbSTATUS.Panels(X).Key, Len(sbSTATUS.Panels(X).Key) - 1))
        If IsWindow(hwnd) Then
            ShowWindow hwnd, SW_NORMAL
            SetParent hwnd, 0
            AddCaptionBar hwnd
        End If
    Next
End Sub
Private Sub mnuGET_Click()
    On Error Resume Next
    ListWindows
End Sub
Public Sub AssignChild(hwnd As Long, wCap As String)
    On Error GoTo ErrorAssignChild
    Dim p As Panel
    SetParent hwnd, Me.hwnd
    ShowWindow hwnd, SW_MAXIMIZE
    RemoveCaptionBar hwnd
    Set p = Me.sbSTATUS.Panels.Add(, hwnd & "K", wCap, , Me.imgCLOSE.Picture)
    p.Bevel = sbrRaised
    p.AutoSize = sbrSpring
    Exit Sub
ErrorAssignChild:
    Exit Sub
End Sub

Private Sub sbSTATUS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim p As Panel, hwnd As Long
    If Button = vbLeftButton Then
        For Each p In sbSTATUS.Panels
            If X >= p.Left And X < p.Left + Me.imgCLOSE.Width Then
                hwnd = CLng(Left(p.Key, Len(p.Key) - 1))
                CloseWindowByhWnd hwnd
                Exit For
            End If
        Next
    Else
        Me.PopupMenu mnuFILE
    End If
End Sub

Private Sub sbSTATUS_PanelClick(ByVal Panel As MSComctlLib.Panel)
    On Error Resume Next
    Dim p As Panel, hwnd As Long
    For Each p In Me.sbSTATUS.Panels
        If Panel.Index = p.Index Then
            p.Bevel = sbrInset
            hwnd = CLng(Left(p.Key, Len(p.Key) - 1))
            If IsIconic(hwnd) Then ShowWindow hwnd, SW_NORMAL
            SetForegroundWindow hwnd
        Else
            p.Bevel = sbrRaised
        End If
    Next
End Sub
Private Sub Timer_Timer()
    On Error Resume Next
    Dim p As Panel, hwnd As Long, X As Long
    Dim l As Long, wstr As String * 255, wCap As String
    For X = sbSTATUS.Panels.Count To 1 Step -1
        Set p = sbSTATUS.Panels(X)
        hwnd = CLng(Left(p.Key, Len(p.Key) - 1))
        If IsWindow(hwnd) = False Then
            sbSTATUS.Panels.Remove p.Index
        Else
            l = GetWindowText(hwnd, wstr, Len(wstr) - 1)
            If l <> 0 Then
                wCap = Left(wstr, l)
                p.Text = wCap
            End If
            If hwnd = GetForegroundWindow Then
                p.Bevel = sbrInset
            Else
                p.Bevel = sbrRaised
            End If
            If IsZoomed(hwnd) Then
                MoveWindow hwnd, 0, ScaleY(Me.sbSTATUS.Height, vbTwips, vbPixels), ScaleX(Me.ScaleWidth, vbTwips, vbPixels), ScaleY(Me.ScaleHeight - Me.sbSTATUS.Height, vbTwips, vbPixels), True
            ElseIf IsIconic(hwnd) Then
                ShowWindow hwnd, SW_HIDE
            End If
        End If
    Next
End Sub
Private Sub tmrMONITOR_Timer()
    On Error Resume Next
    Dim X As Long, p As Panel
    Dim l As Long, wstr As String * 255, wCap As String
    X = 0
    Set p = Nothing
    If Me.MonitorFor <> "" Then
        X = FindWindowWild(Me.MonitorFor, False)
        If X <> 0 Then
            Set p = sbSTATUS.Panels(X & "K")
            If p Is Nothing Then
                l = GetWindowText(X, wstr, Len(wstr) - 1)
                If l <> 0 Then
                    wCap = Left(wstr, l)
                    AssignChild X, wCap
                End If
            End If
        End If
    End If
End Sub
