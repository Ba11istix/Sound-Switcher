Attribute VB_Name = "mStartup"
Option Explicit

Public Declare Function PropVariantToVariant Lib "propsys" (ByRef propvar As Any, ByRef var As Variant) As Long

Private Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal PV As Long)

Public Function LPWSTRtoStr(lPtr As Long, Optional ByVal fFree As Boolean = True) As String
    SysReAllocString VarPtr(LPWSTRtoStr), lPtr
    If fFree Then
        Call CoTaskMemFree(lPtr)
    End If
End Function

Public Function IID_IAudioEndpointVolume() As UUID
    Static iid As UUID
    If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5CDF2C82, CInt(&H841E), CInt(&H4546), &H97, &H22, &HC, &HF7, &H40, &H78, &H22, &H9A)
    IID_IAudioEndpointVolume = iid
End Function

Private Sub DEFINE_UUID(Name As UUID, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
    With Name
        .Data1 = L
        .Data2 = w1
        .Data3 = w2
        .Data4(0) = B0
        .Data4(1) = b1
        .Data4(2) = b2
        .Data4(3) = B3
        .Data4(4) = b4
        .Data4(5) = b5
        .Data4(6) = b6
        .Data4(7) = b7
    End With
End Sub

Public Function UUID_NULL() As UUID
    Static bSet As Boolean
    Static iid As UUID
    If bSet = False Then
        With iid
            .Data1 = 0
            .Data2 = 0
            .Data3 = 0
            .Data4(0) = 0
            .Data4(1) = 0
            .Data4(2) = 0
            .Data4(3) = 0
            .Data4(4) = 0
            .Data4(5) = 0
            .Data4(6) = 0
            .Data4(7) = 0
        End With
    End If
    bSet = True
    UUID_NULL = iid
End Function

Public Function PKEY_Device_FriendlyName() As PROPERTYKEY
    Static pkk As PROPERTYKEY
    If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 14)
    PKEY_Device_FriendlyName = pkk
End Function

Private Sub DEFINE_PROPERTYKEY(Name As PROPERTYKEY, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte, pid As Long)
    With Name.fmtid
        .Data1 = L
        .Data2 = w1
        .Data3 = w2
        .Data4(0) = B0
        .Data4(1) = b1
        .Data4(2) = b2
        .Data4(3) = B3
        .Data4(4) = b4
        .Data4(5) = b5
        .Data4(6) = b6
        .Data4(7) = b7
    End With
    Name.pid = pid
End Sub

Sub Main()
    Load frmMain
End Sub

Public Function InIDE(Optional ByRef B As Boolean = True) As Boolean
    If B = True Then Debug.Assert Not InIDE(InIDE) Else B = True
End Function
