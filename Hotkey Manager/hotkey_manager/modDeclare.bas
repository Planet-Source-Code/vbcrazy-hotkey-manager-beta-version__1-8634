Attribute VB_Name = "modDeclare"
Option Explicit

Public Cn As Connection

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Function GetWindowFromPoint(hwnd As Long, X As Long, Y As Long) As Long
    Dim pt As POINTAPI
    Dim foundhWnd As Long
    pt.X = X
    pt.Y = Y
    foundhWnd = WindowFromPoint(pt.X, pt.Y)
    GetWindowFromPoint = foundhWnd
End Function

