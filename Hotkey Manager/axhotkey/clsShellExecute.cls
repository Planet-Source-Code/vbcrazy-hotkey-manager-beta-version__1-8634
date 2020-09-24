VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShellExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10

Public Sub RunShellExecute(sTopic As String, sFile As Variant, _
    sParams As Variant, sDirectory As Variant, _
    nShowCmd As Long)
    Dim hWndDesk As Long
    Dim success As Long
    Const SE_ERR_NOASSOC = &H31
    'the desktop will be the default for error messages
    hWndDesk = GetDesktopWindow()
    'execute the passed operation
    success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)
End Sub



