VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sound Switcher"
   ClientHeight    =   4455
   ClientLeft      =   6750
   ClientTop       =   5070
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6615
   Begin VB.Timer tHotKey 
      Interval        =   10
      Left            =   6120
      Top             =   120
   End
   Begin VB.CommandButton cmdDud 
      Caption         =   "NotUsed"
      Height          =   195
      Left            =   5640
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5040
      Width           =   735
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   20
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox txtMicVolCOM 
      Height          =   285
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   15
      Text            =   "txtMicVol"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtSpkVolCOM 
      Height          =   285
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   11
      Text            =   "txtSpkVol"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtMicVol 
      Height          =   285
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   13
      Text            =   "txtMicVol"
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtTag 
      Height          =   285
      Left            =   2640
      MaxLength       =   35
      TabIndex        =   17
      Text            =   "txtTag"
      Top             =   4080
      Width           =   2775
   End
   Begin VB.CommandButton cmdActivate 
      Caption         =   "Activate"
      Height          =   615
      Left            =   4800
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   615
      Left            =   4800
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.CheckBox chkCOMMic 
      Caption         =   "Communications"
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CheckBox chkMMMic 
      Caption         =   "Multimedia"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CheckBox chkCOMSpe 
      Caption         =   "Communications"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CheckBox chkMMSpe 
      Caption         =   "Multimedia"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CheckBox chkStart 
      Caption         =   "Use as start up"
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdSavePreset 
      Caption         =   "Save"
      Height          =   375
      Left            =   5520
      TabIndex        =   18
      Top             =   4035
      Width           =   975
   End
   Begin VB.ListBox lstPreset 
      Height          =   1755
      IntegralHeight  =   0   'False
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.PictureBox pSysTray 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   0
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstMic 
      Height          =   1035
      Left            =   3360
      TabIndex        =   7
      Top             =   2040
      Width           =   3135
   End
   Begin VB.ListBox lstSpeaker 
      Height          =   1035
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox txtSpkVol 
      Height          =   285
      Left            =   360
      MaxLength       =   4
      TabIndex        =   9
      Text            =   "txtSpkVol"
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblMicVolCOM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   25
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lblMicVol 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   24
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lblSpkVolCOM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   23
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lblSpkVol 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   21
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lblTag 
      Alignment       =   1  'Right Justify
      Caption         =   "Tag"
      Height          =   255
      Left            =   2160
      TabIndex        =   19
      Top             =   4125
      Width           =   375
   End
   Begin VB.Menu mTrayRight 
      Caption         =   "mTrayRight"
      Visible         =   0   'False
      Begin VB.Menu mChangeSettings 
         Caption         =   "Change Settings"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mFormRight 
      Caption         =   "mFormRight"
      Visible         =   0   'False
      Begin VB.Menu mStartTray 
         Caption         =   "Start in tray"
      End
      Begin VB.Menu mResetLoop 
         Caption         =   "Activate resets loop"
         Checked         =   -1  'True
      End
      Begin VB.Menu mMinimise 
         Caption         =   "Minimise to tray"
      End
      Begin VB.Menu mClose 
         Caption         =   "Close to tray"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Resources: APPICON, ICON / PIC 0 (Speaker), ICON / PIC 1 (Headphones)
Option Explicit

Private Const MOD_ALT = &H1 'Alt
Private Const MOD_CONTROL = &H2 'Ctrl
Private Const MOD_SHIFT = &H4 'Shift
Private Const MOD_WIN = &H8 'Windows key

Private Const PM_REMOVE = &H1
Private Const WM_HOTKEY = &H312

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private WithEvents SysTray As cSysTray
Attribute SysTray.VB_VarHelpID = -1

Private speakerSID() As String
Private micSID() As String

Private currentItem As Integer

Private CurrentIconID As Integer

Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long

Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer

Private bCancel As Boolean
Private HotKey As Long

Private Sub LoadPreset()
    Dim i As Integer
    Dim tSetting As String
    lstPreset.Clear
    Do
        tSetting = GetSetting(App.Title, "Settings", CStr(i), vbNullString)
        If tSetting = vbNullString Then Exit Do
        lstPreset.AddItem tSetting

        If GetSetting(App.Title, "Settings\" & tSetting, "Checked", vbNullString) = "1" Then
            lstPreset.Selected(lstPreset.NewIndex) = True
        Else
            lstPreset.Selected(lstPreset.NewIndex) = False
        End If

        If GetSetting(App.Title, "Settings\" & tSetting, "Startup", vbNullString) = "1" Then
            ActivatePreset lstPreset.NewIndex
        End If

        i = i + 1
    Loop
End Sub

Private Sub SavePreset()
    If chkMMSpe.Value = 1 Then
        SaveSetting App.Title, "Settings\" & txtTag.Text, "MMS", speakerSID(lstSpeaker.ItemData(lstSpeaker.ListIndex))
        SaveSetting App.Title, "Settings\" & txtTag.Text, "MMSV", CStr(lblSpkVol.Caption)
    End If

    If chkCOMSpe.Value = 1 Then
        SaveSetting App.Title, "Settings\" & txtTag.Text, "COS", speakerSID(lstSpeaker.ItemData(lstSpeaker.ListIndex))
        SaveSetting App.Title, "Settings\" & txtTag.Text, "COSV", CStr(lblSpkVolCOM.Caption)
    End If

    If chkMMMic.Value = 1 Then
        SaveSetting App.Title, "Settings\" & txtTag.Text, "MMM", micSID(lstMic.ItemData(lstMic.ListIndex))
        SaveSetting App.Title, "Settings\" & txtTag.Text, "MMMV", CStr(lblMicVol.Caption)
    End If

    If chkCOMMic.Value = 1 Then
        SaveSetting App.Title, "Settings\" & txtTag.Text, "COM", micSID(lstMic.ItemData(lstMic.ListIndex))
        SaveSetting App.Title, "Settings\" & txtTag.Text, "COMV", CStr(lblMicVolCOM.Caption)
    End If

    SaveSetting App.Title, "Settings\" & txtTag.Text, "Icon", CStr(CurrentIconID)
    SaveSetting App.Title, "Settings\" & txtTag.Text, "Startup", CStr(Abs(chkStart.Value))
    SaveSetting App.Title, "Settings\" & txtTag.Text, "Checked", "1"

    SaveSetting App.Title, "Settings", lstPreset.ListCount, CStr(txtTag.Text)

    lstPreset.AddItem txtTag.Text
    lstPreset.Selected(lstPreset.NewIndex) = True
End Sub

Private Function CheckSettings() As Boolean
    Dim i As Integer

    CheckSettings = True
    If txtTag.Text = vbNullString Then
        CheckSettings = False
    Else
        For i = 0 To lstPreset.ListCount - 1
            If GetSetting(App.Title, "Settings", i, vbNullString) = txtTag.Text Then CheckSettings = False
        Next i
    End If
    If Not IsNumeric(txtSpkVol.Text) Then CheckSettings = False
    If Not IsNumeric(txtSpkVolCOM.Text) Then CheckSettings = False
    If Not IsNumeric(txtMicVol.Text) Then CheckSettings = False
    If Not IsNumeric(txtMicVolCOM.Text) Then CheckSettings = False
    If chkMMSpe.Value + chkCOMSpe.Value + chkMMMic.Value + chkCOMMic.Value = 0 Then CheckSettings = False
End Function

Private Sub DeletePreset()
    Dim i As Integer
    If lstPreset.ListCount = 0 Or lstPreset.List(lstPreset.ListIndex) = vbNullString Then Exit Sub
    If MsgBox("Are you sure that you want to delete the " & lstPreset.List(lstPreset.ListIndex) & " preset?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
        For i = 0 To lstPreset.ListCount - 1    'delete these first - should mean if user restarts after crash error is fixed...
            DeleteSetting App.Title, "Settings", i
        Next i
        DeleteSetting App.Title, "Settings\" & lstPreset.List(lstPreset.ListIndex)

        lstPreset.RemoveItem lstPreset.ListIndex
        If lstPreset.ListCount = 0 Then Exit Sub
        For i = 0 To lstPreset.ListCount - 1
            SaveSetting App.Title, "Settings", i, lstPreset.List(i)
        Next i
    End If
End Sub

Private Sub ActivatePreset(presetID As Integer)
    If presetID = -1 Then Exit Sub
    lstPreset.ListIndex = presetID
    SetDefault True, GetSetting(App.Title, "Settings\" & lstPreset.List(presetID), "MMS", vbNullString), CSng(GetSetting(App.Title, "Settings\" & lstPreset.List(presetID), "MMSV", -1)), GetSetting(App.Title, "Settings\" & lstPreset.List(presetID), "COS", vbNullString), CSng(GetSetting(App.Title, "Settings\" & lstPreset.List(presetID), "COSV", -1))
    SetDefault False, GetSetting(App.Title, "Settings\" & lstPreset.List(presetID), "MMM", vbNullString), CSng(GetSetting(App.Title, "Settings\" & lstPreset.List(presetID), "MMMV", -1)), GetSetting(App.Title, "Settings\" & lstPreset.List(presetID), "COM", vbNullString), CSng(GetSetting(App.Title, "Settings\" & lstPreset.List(presetID), "COMV", -1))
    SysTray.SetIcon "ICON" & CInt(GetSetting(App.Title, "Settings\" & lstPreset.List(presetID), "Icon", vbNullString))
    SysTray.TooltipText = lstPreset.List(presetID)
End Sub

Private Function GetNextPreset() As Integer
    Dim i As Integer
    If lstPreset.ListCount = 0 Then
        GetNextPreset = -1
        Exit Function
    End If
    i = currentItem
    Do
        i = i + 1
        If i > lstPreset.ListCount - 1 Then i = 0
        If lstPreset.Selected(i) Then Exit Do
    Loop Until i = currentItem
    currentItem = i
    GetNextPreset = i
End Function

Private Sub cmdActivate_Click()
    If mResetLoop.Checked = True Then currentItem = lstPreset.ListIndex
    ActivatePreset lstPreset.ListIndex
End Sub

Private Sub cmdDelete_Click()
    DeletePreset
End Sub

Private Sub cmdMoveDown_Click()
    MoveListItem lstPreset.ListIndex, lstPreset.ListIndex + 1
End Sub

Private Sub cmdMoveUp_Click()
    MoveListItem lstPreset.ListIndex, lstPreset.ListIndex - 1
End Sub

Private Function MoveListItem(CurrentIndex As Integer, NewIndex As Integer)
    Dim tItem As String
    Dim tSelected As Boolean
    Dim tData As String
    Dim tReg As String

    With lstPreset
        If .ListCount = NewIndex Or NewIndex = -1 Or NewIndex = -2 Then Exit Function

        tItem = .List(CurrentIndex)
        tSelected = .Selected(CurrentIndex)
        tData = .ItemData(CurrentIndex)

        .RemoveItem CurrentIndex
        .AddItem tItem, NewIndex

        .ListIndex = NewIndex
        .Selected(NewIndex) = tSelected
        .ItemData(NewIndex) = tData

        tReg = GetSetting(App.Title, "Settings", CStr(CurrentIndex))
        SaveSetting App.Title, "Settings", CStr(CurrentIndex), GetSetting(App.Title, "Settings", CStr(NewIndex))
        SaveSetting App.Title, "Settings", CStr(NewIndex), tReg
    End With
End Function

Private Sub RefreshSpeaker()
    Dim i As Integer
    speakerSID = GetDevices(eRender)
    lstSpeaker.Clear
    For i = 0 To UBound(speakerSID) - 1
        lstSpeaker.AddItem GetDeviceName(speakerSID(i))
        lstSpeaker.ItemData(lstSpeaker.NewIndex) = i
    Next i
End Sub

Private Sub RefreshMic()
    Dim i As Integer
    micSID = GetDevices(eCapture)
    lstMic.Clear
    For i = 0 To UBound(micSID) - 1
        lstMic.AddItem GetDeviceName(micSID(i))
        lstMic.ItemData(lstMic.NewIndex) = i
    Next i
End Sub

Private Sub cmdSavePreset_Click()
    If Not CheckSettings Then Exit Sub
    txtTag.Text = Trim(txtTag.Text)
    SavePreset
    ResetForm
End Sub

Private Sub Form_Click()
    cmdDud.SetFocus
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Set SysTray = New cSysTray
    SysTray.Init pSysTray.hWnd, Me.Icon, Me.hWnd
    SysTray.SetIcon "APPICON"
    SysTray.TooltipText = "Sound Switcher"
    
    RefreshSpeaker
    RefreshMic
    
    LoadNextIcon 0

    LoadPreset

    ResetForm
    
    HotKey = GlobalAddAtom("NextSoundPreset")
    If RegisterHotKey(Me.hWnd, HotKey, MOD_CONTROL, vbKeyM) Then tHotKey.Enabled = True
    
    i = CInt(GetSetting(App.Title, "Settings", "StartInTray", -1))
    If i = -1 Then
        SaveSetting App.Title, "Settings", "StartInTray", "0"
        Me.Show
    Else
        mStartTray.Checked = i
        If i <> 0 Then
            SysTray.ShowWindow Me.hWnd, False
        Else
            Me.Show
        End If
    End If

    i = CInt(GetSetting(App.Title, "Settings", "ActivateReset", -1))
    If i = -1 Then
        SaveSetting App.Title, "Settings", "ActivateReset", "1"
    Else
        mResetLoop.Checked = i
    End If

    i = CInt(GetSetting(App.Title, "Settings", "CloseToTray", -1))
    If i = -1 Then
        SaveSetting App.Title, "Settings", "CloseToTray", "1"
    Else
        mClose.Checked = i
    End If

    i = CInt(GetSetting(App.Title, "Settings", "MinimiseToTray", -1))
    If i = -1 Then
        SaveSetting App.Title, "Settings", "MinimiseToTray", "0"
    Else
        mMinimise.Checked = i
    End If
End Sub

Private Sub ResetForm()
    txtTag.Text = vbNullString
    txtSpkVol.Text = "1.00"
    txtSpkVolCOM.Text = "1.00"
    txtMicVol.Text = "1.00"
    txtMicVolCOM.Text = "1.00"
    chkMMSpe.Value = 1
    chkMMMic.Value = 1
    chkCOMSpe.Value = 1
    chkCOMMic.Value = 1
    chkStart.Value = 0

    If lstSpeaker.ListCount > 0 Then
        lstSpeaker.ListIndex = 0
    Else
        cmdSavePreset.Enabled = False
    End If
    If lstMic.ListCount > 0 Then
        lstMic.ListIndex = 0
    Else
        cmdSavePreset.Enabled = False
    End If
End Sub

Private Sub LoadNextIcon(Optional iconID As Integer = -1)
    If iconID <> -1 Then
        CurrentIconID = iconID - 1
    End If

    On Error Resume Next: Do
        CurrentIconID = CurrentIconID + 1
        Err.Clear
        picIcon.Picture = LoadResPicture("PIC" & CurrentIconID, vbResBitmap)
        If CurrentIconID > 100 Or Err.Number = 326 Then CurrentIconID = -1
    Loop Until Err.Number = 0: On Error GoTo 0
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mFormRight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 And mClose.Checked = True Then
        SysTray.ShowWindow Me.hWnd, False
        Cancel = True
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized And mMinimise.Checked = True Then
        Me.Visible = False
        Me.WindowState = vbNormal
        Me.Visible = True
        SysTray.ShowWindow Me.hWnd, False
    End If
End Sub

Private Sub lblMicVol_Click(): txtMicVol.SetFocus: End Sub
Private Sub lblMicVolCOM_Click(): txtMicVolCOM.SetFocus: End Sub
Private Sub lblSpkVol_Click(): txtSpkVol.SetFocus: End Sub
Private Sub lblSpkVolCOM_Click(): txtSpkVolCOM.SetFocus: End Sub

Private Sub lstPreset_ItemCheck(Item As Integer)
    If lstPreset.ListCount = 0 Then Exit Sub
    SaveSetting App.Title, "Settings\" & lstPreset.List(Item), "Checked", CStr(Abs(lstPreset.Selected(Item)))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SysTray = Nothing
    bCancel = True
    GlobalDeleteAtom HotKey
    If tHotKey.Enabled Then UnregisterHotKey Me.hWnd, HotKey
End Sub

Private Sub mChangeSettings_Click(): Me.WindowState = vbNormal: SysTray.ShowWindow Me.hWnd: End Sub
Private Sub mExit_Click(): Unload Me: End Sub

Private Sub mClose_Click()
    mClose.Checked = Not mClose.Checked
    SaveSetting App.Title, "Settings", "CloseToTray", CStr(Abs(mClose.Checked))
End Sub

Private Sub mMinimise_Click()
    mMinimise.Checked = Not mMinimise.Checked
    SaveSetting App.Title, "Settings", "MinimiseToTray", CStr(Abs(mMinimise.Checked))
End Sub

Private Sub mResetLoop_Click()
    mResetLoop.Checked = Not mResetLoop.Checked
    SaveSetting App.Title, "Settings", "ActivateReset", CStr(Abs(mResetLoop.Checked))
End Sub

Private Sub mStartTray_Click()
    mStartTray.Checked = Not mStartTray.Checked
    SaveSetting App.Title, "Settings", "StartInTray", CStr(Abs(mStartTray.Checked))
End Sub

Private Sub picIcon_Click()
    LoadNextIcon
End Sub

Private Sub pSysTray_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single):
    SysTray.TooltipText = lstPreset.List(currentItem) & ": " & GetVolume
    SysTray.MouseMove x
End Sub

Private Sub SysTray_LeftClick(): ActivatePreset GetNextPreset: End Sub
Private Sub SysTray_RightClick(): PopupMenu mTrayRight: End Sub

Private Sub SetDefault(Speaker As Boolean, Multimedia As String, Optional Volume As Single = -1, Optional Communication As String = vbNullString, Optional comVolume As Single = -1)
    Dim pCfg As PolicyConfigClient
    Dim volu As IAudioEndpointVolume
    Dim devi As IMMDevice
    Dim devEnum As MMDeviceEnumerator
    Dim sType As EDataFlow

    If Speaker = True Then
        sType = eRender
    Else
        sType = eCapture
    End If

    Set pCfg = New PolicyConfigClient
    On Error Resume Next
    pCfg.SetDefaultEndpoint StrPtr(Multimedia), eMultimedia

    Set devEnum = New MMDeviceEnumerator
    If Volume >= 0 Then
        devEnum.GetDefaultAudioEndpoint sType, eMultimedia, devi
        devi.Activate IID_IAudioEndpointVolume, CLSCTX_INPROC_SERVER, CVar(10), volu
        volu.SetMute 0, UUID_NULL
        volu.SetMasterVolumeLevelScalar Volume, UUID_NULL
    End If

    If Communication <> vbNullString Then
        pCfg.SetDefaultEndpoint StrPtr(Communication), eCommunications
        If comVolume >= 0 Then
            devEnum.GetDefaultAudioEndpoint sType, eCommunications, devi
            devi.Activate IID_IAudioEndpointVolume, CLSCTX_INPROC_SERVER, CVar(10), volu
            volu.SetMute 0, UUID_NULL
            volu.SetMasterVolumeLevelScalar comVolume, UUID_NULL
        End If
    End If
    Set devEnum = Nothing
    Set pCfg = Nothing
End Sub

Private Function GetDevices(tType As EDataFlow) As String()
    Dim sTemp() As String
    Dim i As Long
    Dim nCount As Long
    Dim lpID As Long
    Dim pDvCol As IMMDeviceCollection
    Dim pDevice As IMMDevice
    Dim pDvEnum As MMDeviceEnumerator

    Set pDvEnum = New MMDeviceEnumerator
    pDvEnum.EnumAudioEndpoints tType, DEVICE_STATE_ACTIVE, pDvCol
    Set pDvEnum = Nothing

    pDvCol.GetCount nCount
    ReDim sTemp(nCount) As String
    For i = 0 To (nCount - 1)
        pDvCol.Item i, pDevice
        pDevice.GetId lpID
        sTemp(i) = LPWSTRtoStr(lpID)
    Next
    GetDevices = sTemp
End Function

Private Function GetVolume() As String
    Dim vol As Single
    Dim volu As IAudioEndpointVolume
    Dim devi As IMMDevice
    Dim devEnum As MMDeviceEnumerator
    
    Dim mute As BOOL
    
    Set devEnum = New MMDeviceEnumerator
    devEnum.GetDefaultAudioEndpoint eRender, eMultimedia, devi
    Set devEnum = Nothing
    devi.Activate IID_IAudioEndpointVolume, CLSCTX_INPROC_SERVER, CVar(10), volu

    volu.GetMute mute
    If mute Then
        GetVolume = "Muted"
    Else
        volu.GetMasterVolumeLevelScalar vol
        GetVolume = CStr(CInt(vol * 100)) & "%"
    End If
End Function

Private Function GetDeviceName(SID As String) As String
    Dim pDevice As IMMDevice
    Dim pStore As IPropertyStore
    Dim vrProp As Variant
    Dim vProp As Variant
    Dim vte As VbVarType
    Dim devEnum As MMDeviceEnumerator

    Set devEnum = New MMDeviceEnumerator
    devEnum.GetDevice StrPtr(SID), pDevice
    Set devEnum = Nothing

    pDevice.OpenPropertyStore STGM_READ, pStore
    If Not pStore Is Nothing Then
        pStore.GetValue PKEY_Device_FriendlyName, vProp
        PropVariantToVariant vProp, vrProp

        vte = VarType(vrProp)
        Select Case vte
        Case vbDataObject, vbObject, vbUserDefinedType
            GetDeviceName = "<cannot display this type>"
        Case vbEmpty, vbNull
            GetDeviceName = "<empty or null>"
        Case vbError
            GetDeviceName = "<vbError>"
        Case Else
            GetDeviceName = CStr(vrProp)
        End Select
    End If
End Function

Private Sub tHotKey_Timer()
    Dim Message As Msg
    WaitMessage
    If PeekMessage(Message, Me.hWnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then
        Select Case Message.wParam
            Case HotKey
                ActivatePreset GetNextPreset
        End Select
    End If
End Sub

Private Sub txtMicVol_Change(): lblMicVol.Caption = txtMicVol.Text: End Sub
Private Sub txtMicVol_GotFocus(): SelectAll txtMicVol: lblMicVol.BackColor = &H8000000D: End Sub
Private Sub txtMicVol_KeyPress(KeyAscii As Integer): CheckAscii KeyAscii: End Sub
Private Sub txtMicVol_LostFocus(): lblMicVol.BackColor = &H8000000F: FixCaption lblMicVol: End Sub

Private Sub txtMicVolCOM_Change(): lblMicVolCOM.Caption = txtMicVolCOM.Text: End Sub
Private Sub txtMicVolCOM_GotFocus(): SelectAll txtMicVolCOM: lblMicVolCOM.BackColor = &H8000000D: End Sub
Private Sub txtMicVolCOM_KeyPress(KeyAscii As Integer): CheckAscii KeyAscii: End Sub
Private Sub txtMicVolCOM_LostFocus(): lblMicVolCOM.BackColor = &H8000000F: FixCaption lblMicVolCOM: End Sub

Private Sub txtSpkVol_Change(): lblSpkVol.Caption = txtSpkVol.Text: End Sub
Private Sub txtSpkVol_GotFocus(): SelectAll txtSpkVol: lblSpkVol.BackColor = &H8000000D: End Sub
Private Sub txtSpkVol_KeyPress(KeyAscii As Integer): CheckAscii KeyAscii: End Sub
Private Sub txtSpkVol_LostFocus(): lblSpkVol.BackColor = &H8000000F: FixCaption lblSpkVol: End Sub

Private Sub txtSpkVolCOM_Change(): lblSpkVolCOM.Caption = txtSpkVolCOM.Text: End Sub
Private Sub txtSpkVolCOM_GotFocus(): SelectAll txtSpkVolCOM: lblSpkVolCOM.BackColor = &H8000000D: End Sub
Private Sub txtSpkVolCOM_KeyPress(KeyAscii As Integer): CheckAscii KeyAscii: End Sub
Private Sub txtSpkVolCOM_LostFocus(): lblSpkVolCOM.BackColor = &H8000000F: FixCaption lblSpkVolCOM: End Sub

Private Sub CheckAscii(ByRef KeyAscii As Integer)    'is number, dot, backspace or delete
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 127 Then KeyAscii = 0
End Sub

Private Sub FixCaption(ByRef lbl As Label)
    With lbl
        If Not IsNumeric(.Caption) Then
            .Caption = "0.00"
        ElseIf CDbl(.Caption) > 1 Then
            .Caption = "1.00"
        Else
            .Caption = Format(.Caption, "0.00")
        End If
    End With
End Sub

Private Sub SelectAll(ByRef tTextBox As TextBox)
    With tTextBox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTag_KeyPress(KeyAscii As Integer)    'is alphanumeric, backspace or delete
    If (KeyAscii < 48 Or (KeyAscii > 57 And KeyAscii < 65) Or (KeyAscii > 90 And KeyAscii < 97) Or KeyAscii > 122) And KeyAscii <> 127 And KeyAscii <> 8 And KeyAscii <> 32 Then    'alphanumeric and space only
        KeyAscii = 0
    End If
End Sub
