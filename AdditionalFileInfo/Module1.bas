Attribute VB_Name = "Module1"
Option Explicit
Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As Long, ByVal lpType As Long) As Long
Declare Function FindResourceEx Lib "kernel32" Alias "FindResourceExA" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal wLanguage As Integer) As Long
Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
'VERSION
Public Type VS_FIXEDFILEINFO
    dwSignature As Long  'Contains the value 0xFEEFO4BD. This is used with the szKey member of VS_VERSION_INFO data when searching a file for the VS_FIXEDFILEINFO structure.
    dwStrucVersion As Long ' e.g. 0x00000042 = "0.42"
    dwFileVersionMS As Long ' e.g. 0x00030075 = "3.75"
    dwFileVersionLS As Long ' e.g. 0x00000031 = "0.31"
    dwProductVersionMS As Long ' e.g. 0x00030010 = "3.10"
    dwProductVersionLS As Long ' e.g. 0x00000031 = "0.31"
    dwFileFlagsMask As Long ' = 0x3F for version "0.42"
    dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
    dwFileType As Long ' e.g. VFT_DRIVER
    dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long ' e.g. 0
    dwFileDateLS As Long ' e.g. 0
End Type

Public Type INF_INF '(Version Info)
 wLength As Integer
 wValueLength As Integer
 wType As Integer
End Type

Dim SINF2 As String
Dim Vlng As Long
Dim VVal As Integer
Dim check1 As Byte
Dim totalLG As Long
Dim TcontX As Long
Dim llng As Long
Dim SINF As String
Dim tmpINF As INF_INF
Private FFInfo As VS_FIXEDFILEINFO
Private ResTotLen As Long
Private OINF As String
Private Sub GetFileInfo(data() As Byte)
Dim countX As Long
Dim tcountX As Long

CopyMemory tmpINF, data(0), Len(tmpINF)
countX = countX + Len(tmpINF)

Dim FINF As String
llng = lstrlenW(ByVal VarPtr(data(countX)))
FINF = Space(llng)
CopyMemory ByVal StrPtr(FINF), data(countX), llng * 2
countX = countX + llng * 2 + 2
If (countX Mod 4) <> 0 Then countX = countX + 2


OINF = "Length:" & tmpINF.wLength & " (" & Hex(tmpINF.wLength) & "h) Bytes" & vbCrLf
If tmpINF.wType = 1 Then
OINF = OINF & "Text Resource Version" & vbCrLf
Else
OINF = OINF & "Binary Resource Version" & vbCrLf
End If
OINF = OINF & FINF & vbCrLf

If Not (Not CBool(tmpINF.wValueLength)) Then
CopyMemory FFInfo, data(countX), Len(FFInfo)
countX = countX + Len(FFInfo)
OINF = OINF & "Signature:" & Hex(FFInfo.dwSignature) & "h" & vbCrLf
OINF = OINF & "File Version:" & FormatVER(FFInfo.dwFileVersionMS, FFInfo.dwFileVersionLS) & vbCrLf
OINF = OINF & "Product Version:" & FormatVER(FFInfo.dwProductVersionMS, FFInfo.dwProductVersionLS) & vbCrLf
OINF = OINF & "File OS:" & GetOS(FFInfo.dwFileOS) & vbCrLf
OINF = OINF & "File Type:" & GetFileType(FFInfo.dwFileType) & vbCrLf
End If
OINF = OINF & vbCrLf

Do While countX < ResTotLen
CopyMemory tmpINF, data(countX), Len(tmpINF)

totalLG = countX + tmpINF.wLength 'Ukupna duzina INFO-a

countX = countX + Len(tmpINF)

If tmpINF.wLength = 0 And tmpINF.wType = 0 And tmpINF.wValueLength = 0 Then GoTo eend

tcountX = countX
llng = lstrlenW(ByVal VarPtr(data(countX)))
SINF = Space(llng)
CopyMemory ByVal StrPtr(SINF), data(countX), llng * 2

OINF = OINF & SINF & ":" & vbCrLf
If SINF = "StringFileInfo" Then
countX = countX + llng * 2 + 2
If (countX Mod 4) <> 0 Then countX = countX + 2
countX = GetStrFilInf(data, countX, totalLG)

ElseIf SINF = "VarFileInfo" Then
countX = countX + llng * 2 + 2
If (countX Mod 4) <> 0 Then countX = countX + 2
countX = GetVarFilInf(data, countX)
End If

If (countX Mod 4) <> 0 Then countX = countX + 2
eend:
Loop
End Sub
Private Function GetVarFilInf(data() As Byte, ByVal countX As Long) As Long
Dim u As Long

CopyMemory tmpINF, data(countX), Len(tmpINF)
countX = countX + Len(tmpINF)
llng = lstrlenW(ByVal VarPtr(data(countX)))
SINF = Space(llng)
CopyMemory ByVal StrPtr(SINF), data(countX), llng * 2
OINF = OINF & SINF & ":"
countX = countX + llng * 2 + 2
If (countX Mod 4) <> 0 Then countX = countX + 2

Vlng = tmpINF.wValueLength / 2
For u = 1 To Vlng
CopyMemory VVal, data(countX), 2
OINF = OINF & " " & Hex(VVal) & "h"
countX = countX + 2
Next u
OINF = OINF & vbCrLf
If (countX Mod 4) <> 0 Then countX = countX + 2
GetVarFilInf = countX
OINF = OINF & vbCrLf
End Function


Private Function GetStrFilInf(data() As Byte, ByVal countX As Long, ByVal length As Long) As Long
Dim tcountX As Long

CopyMemory tmpINF, data(countX), Len(tmpINF)
countX = countX + Len(tmpINF)
If tmpINF.wLength = 0 And tmpINF.wType = 0 And tmpINF.wValueLength = 0 Then GoTo Dalje
llng = lstrlenW(ByVal VarPtr(data(countX)))
SINF = Space(llng)
CopyMemory ByVal StrPtr(SINF), data(countX), llng * 2
OINF = OINF & SINF & vbCrLf
countX = countX + llng * 2 + 2
If (countX Mod 4) <> 0 Then countX = countX + 2
Dalje:
Do While countX < length
CopyMemory tmpINF, data(countX), Len(tmpINF)
tcountX = countX + tmpINF.wLength
countX = countX + Len(tmpINF)
llng = lstrlenW(ByVal VarPtr(data(countX)))
SINF = Space(llng)
CopyMemory ByVal StrPtr(SINF), data(countX), llng * 2
OINF = OINF & SINF
countX = countX + llng * 2 + 2
If (countX Mod 4) <> 0 Then countX = countX + 2
If countX = tcountX Then OINF = OINF & ":" & vbCrLf: GoTo nemadalje
llng = lstrlenW(ByVal VarPtr(data(countX)))
SINF = Space(llng)
CopyMemory ByVal StrPtr(SINF), data(countX), llng * 2
OINF = OINF & ":" & SINF & vbCrLf
nemadalje:
countX = tcountX
If (countX Mod 4) <> 0 Then countX = countX + 2
Loop
OINF = OINF & vbCrLf
GetStrFilInf = length
End Function
Private Function FormatVER(ByVal Hvalue As Long, ByVal Lvalue As Long) As String
FormatVER = (Hvalue And &HFFFF0000) / &H10000
FormatVER = FormatVER & "." & (Hvalue And &HFFFF&) & "."
FormatVER = FormatVER & (Lvalue And &HFFFF0000) / &H10000 & "."
FormatVER = FormatVER & (Lvalue And &HFFFF&)
End Function

Private Function GetOS(ByVal value As Long) As String
If value = 0 Then GetOS = "Unknow": Exit Function
If (value And &H10000) = &H10000 Then GetOS = GetOS & "Dos_"
If (value And &H1&) = &H1& Then GetOS = GetOS & "Windows16_"
If (value And &H4&) = &H4& Then GetOS = GetOS & "Windows32_"
If (value And &H40000) = &H40000 Then GetOS = GetOS & "NT_"
If value = &H20000 Then GetOS = "OS/2-16_"
If value = &H20002 Then GetOS = "OS/2-16_PM16_"
If value = &H30000 Then GetOS = "OS/2-32_"
If value = &H30002 Then GetOS = "OS/2-32_PM32_"
GetOS = Left(GetOS, Len(GetOS) - 1)
End Function

Private Function GetFileType(ByVal value As Long) As String
Select Case value
Case 1
GetFileType = "Application"
Case 2
GetFileType = "DLL (Dynamic Link Library)"
Case 3
GetFileType = "Driver"
Case 4
GetFileType = "Font"
Case 5
GetFileType = "Virtual Device"
Case 7
GetFileType = "SLL (Static Link Library)"
Case 0
GetFileType = "Unknow"
End Select
End Function
Public Function InitFileInfo(Optional ByVal Filename As String, Optional ByRef InfoExist As Long) As String
Dim MDLH As Long
Dim FRES As Long
Dim ResI As Long
Dim RPTR As Long
Dim RSZ As Long
Dim INCODE() As Byte

If Len(Filename) = 0 Then
MDLH = App.hInstance
Else
MDLH = LoadLibraryEx(Filename, 0, 2)
End If

ResI = FindResourceEx(MDLH, 16, 1, 0)
If ResI = 0 Or MDLH = 0 Then InfoExist = 0: Exit Function

FRES = LoadResource(MDLH, ResI)
RPTR = LockResource(FRES)
RSZ = SizeofResource(MDLH, ResI)
ReDim INCODE(RSZ - 1)
CopyMemory INCODE(0), ByVal RPTR, RSZ
ResTotLen = RSZ
GetFileInfo INCODE
InitFileInfo = OINF

InfoExist = 1

If Len(Filename) <> 0 Then
FreeLibrary MDLH
End If

End Function
Public Function FindFromStringInfo(ByVal TypeInformation As String) As String
On Error GoTo Dalje
If Len(OINF) = 0 Then Exit Function
Dim SP2() As String
SP2 = Split(OINF, TypeInformation & ":")
Dim SP3() As String
SP3 = Split(SP2(1), vbCrLf)
FindFromStringInfo = SP3(0)
Erase SP2
Erase SP3
Exit Function
Dalje:
On Error GoTo 0
End Function
