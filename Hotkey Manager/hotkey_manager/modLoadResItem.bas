Attribute VB_Name = "modLoadResItem"
Option Explicit

Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Declare Function GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFilename As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function GetTempFile(ByVal strDestPath As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, lpTempFilename As String) As Boolean
    If strDestPath = "" Then
        strDestPath = String(255, vbNullChar)
        If GetTempPath(255, strDestPath) = 0 Then
            GetTempFile = False
            Exit Function
        End If
    End If
    lpTempFilename = String(255, vbNullChar)
    GetTempFile = GetTempFilename(strDestPath, lpPrefixString, wUnique, lpTempFilename) > 0
    lpTempFilename = StripTerminator(lpTempFilename)
End Function

Public Function LoadPictureResource(ByVal ResourceID As Long, ByVal sResourceType As String, ByVal bFlash As Boolean, Optional TempFile) As Picture
    Dim sFileName As String
    If IsMissing(TempFile) Then
        GetTempFile "", "~rs", 0, sFileName
    Else
        sFileName = TempFile
    End If
    If SaveResItemToDisk(ResourceID, sResourceType, sFileName) = 0 Then
        If bFlash Then
            If frmAbout.ShockwaveFlash1.IsPlaying Then
                frmAbout.ShockwaveFlash1.LoadMovie 0, sFileName
            Else
                frmAbout.ShockwaveFlash1.Movie = sFileName
                frmAbout.ShockwaveFlash1.Play
            End If
        Else
            Set LoadPictureResource = LoadPicture(sFileName)
            Kill sFileName
        End If
    End If
End Function

Public Function SaveResItemToDisk(ByVal iResourceNum As Integer, ByVal sResourceType As String, ByVal sDestFileName As String) As Long
    Dim bytResourceData()   As Byte
    Dim iFileNumOut         As Integer
    'On Error GoTo SaveResItemToDisk_err
    bytResourceData = LoadResData(iResourceNum, sResourceType)
    iFileNumOut = FreeFile
    Open sDestFileName For Binary Access Write As #iFileNumOut
        Put #iFileNumOut, , bytResourceData
    Close #iFileNumOut
    SaveResItemToDisk = 0
    Exit Function
SaveResItemToDisk_err:
    SaveResItemToDisk = Err.Number
End Function

Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
