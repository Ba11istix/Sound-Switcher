VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As IconData) As Boolean
Private Declare Function lShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long

Public Event LeftClick()
Public Event RightClick()

Private Type IconData
    iSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeoutOrVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

Public Enum BalloonIconEnum
    beNone = 0
    beInformation = 1
    beWarning = 2
    beError = 3
    beNoSound = &H10
End Enum

Private mIcon As IconData
Private mTip As String

Private Sub SetAppIcon(ByVal hWnd As Long)
    Dim lhWnd As Long
    Dim lhWndTop As Long
    Dim hIcon As Long
    lhWnd = hWnd
    Do Until lhWnd = 0
        lhWndTop = lhWnd
        lhWnd = GetWindow(lhWnd, 4)
    Loop
    If InIDE Then Exit Sub
    On Error GoTo handler    'skip if not exist
    hIcon = LoadImageAsString(App.hInstance, "APPICON", 1, GetSystemMetrics(11), GetSystemMetrics(12), &H8000&)
    SendMessageLong lhWndTop, &H80, 1, hIcon
    SendMessageLong hWnd, &H80, 1, hIcon

    hIcon = LoadImageAsString(App.hInstance, "APPICON", 1, GetSystemMetrics(49), GetSystemMetrics(50), &H8000&)
    SendMessageLong lhWndTop, &H80, 0, hIcon
    SendMessageLong hWnd, &H80, 0, hIcon
handler:
End Sub

Public Sub SetIcon(ByVal ResID As Variant)
    On Error Resume Next    'skip if not exist
    If InIDE Then
        mIcon.hIcon = LoadResPicture(ResID, vbResIcon)
    Else
        mIcon.hIcon = LoadImageAsString(App.hInstance, ResID, 1, GetSystemMetrics(49), GetSystemMetrics(50), &H8000&)
    End If
    RefreshIcon
End Sub

Public Sub Init(ByVal hWnd As Long, ByVal fIcon As Long, Optional ByVal APPhWnd As Long)
    If APPhWnd <> 0 Then SetAppIcon APPhWnd
    With mIcon
        .hWnd = hWnd
        .uFlags = &H1
        .uCallbackMessage = &H200
        .hIcon = fIcon
        .iSize = Len(mIcon)
    End With
    NotifyIcon &H0, mIcon
    RefreshIcon
End Sub

Public Sub ShowWindow(ByVal hWnd As Long, Optional Visible As Boolean = True)
    If Visible Then
        lShowWindow hWnd, 5
    Else
        lShowWindow hWnd, 0
    End If
End Sub

Public Sub ShowBalloonTip(ByVal Message As String, Optional ByVal iIcon As BalloonIconEnum, Optional ByVal Title As String, Optional ByVal Timeout = 5000)
    mIcon.szInfo = Message & vbNullChar
    mIcon.szInfoTitle = Title & " " & vbNullChar
    mIcon.uTimeoutOrVersion = Timeout
    mIcon.dwInfoFlags = iIcon
    mIcon.uFlags = &H10
    NotifyIcon &H1, mIcon
End Sub

Public Sub MouseMove(ByVal x As Single)             '&H200 MouseMove
    Select Case (x / Screen.TwipsPerPixelX)         '&H201 LeftMouseDown
    Case &H201: SetForegroundWindow mIcon.hWnd: RaiseEvent LeftClick                '&H202 LeftMouseUp
    Case &H204: SetForegroundWindow mIcon.hWnd: RaiseEvent RightClick               '&H203 LeftDblClick
    End Select                                      '&H204 RightMouseDown
End Sub                                             '&H205 RightMouseUp
'&H206 RightDblClick
Private Sub Class_Terminate()
    NotifyIcon &H2, mIcon
End Sub

Private Sub RefreshIcon()
    mIcon.uFlags = &H2
    NotifyIcon &H1, mIcon
End Sub

Public Property Get TooltipText() As String
    TooltipText = mTip
End Property

Public Property Let TooltipText(ByVal pstrTooltip As String)
    Const NIF_TIP = &H4
    Const NIM_MODIFY = &H1
    
    mTip = pstrTooltip
    mIcon.szTip = mTip & vbNullChar
    mIcon.uFlags = NIF_TIP
    NotifyIcon NIM_MODIFY, mIcon
End Property
