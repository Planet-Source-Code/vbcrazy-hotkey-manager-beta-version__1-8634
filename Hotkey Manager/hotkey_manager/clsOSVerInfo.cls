VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOSVerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Public Enum TypeOfOS
    VER_PLATFORM_WIN32_WINDOWS = 1
    ver_platform_win32_nt = 2
End Enum

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVersionInfo) As Long

Private Type OSVersionInfo
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Dim tempOSVerInfo As OSVersionInfo

Public Function dwOSVersionInfoSize() As Long
    dwOSVersionInfoSize = tempOSVerInfo.dwOSVersionInfoSize
End Function

Public Function dwMajorVersion() As Long
    dwMajorVersion = tempOSVerInfo.dwMajorVersion
End Function

Public Function dwMinorVersion() As Long
    dwMinorVersion = tempOSVerInfo.dwMinorVersion
End Function

Public Function dwBuildNumber() As Long
    dwBuildNumber = tempOSVerInfo.dwBuildNumber And &HFFFF&
End Function

Public Function dwPlatformId() As Long
    dwPlatformId = tempOSVerInfo.dwPlatformId
End Function

Public Function szCSDVersion() As String
    szCSDVersion = LPSTRToVBString$(tempOSVerInfo.szCSDVersion)
End Function

Public Sub WindowsVersion()
    Dim dl As Long
    tempOSVerInfo.dwOSVersionInfoSize = 148
    dl = GetVersionEx(tempOSVerInfo)
End Sub

Private Function LPSTRToVBString(ByVal strANY As String) As String
    Dim nullpos As Long
    nullpos = InStr(strANY, Chr$(0))
    If nullpos > 0 Then
        LPSTRToVBString = Left$(strANY, nullpos - 1)
    Else
        LPSTRToVBString = ""
    End If
End Function
