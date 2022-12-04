Attribute VB_Name = "modCenterDlg"
Option Explicit

Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hwnd As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Const WH_CALLWNDPROC = 4
Public Const WM_INITDIALOG = &H110
Public Const GWL_WNDPROC = (-4)
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10

Public lWndProc As Long
Public hHook As Long, lHookWndProc As Long

Public Function AppHook(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim CWP As CWPSTRUCT
    CopyMemory CWP, ByVal lParam, Len(CWP)
    Select Case CWP.message
        Case WM_INITDIALOG
            lWndProc = SetWindowLong(CWP.hwnd, GWL_WNDPROC, AddressOf Dlg_WndProc)
            AppHook = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
            UnhookWindowsHookEx hHook
            hHook = 0
            Exit Function
    End Select
    AppHook = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
End Function

Public Function Dlg_WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case Msg
        Case WM_INITDIALOG
            Dim R As RECT, x As Long, y As Long
            GetWindowRect hwnd, R
            x = (frmCustom.Left \ Screen.TwipsPerPixelX + (frmCustom.Width \ Screen.TwipsPerPixelX - (R.Right - R.Left)) \ 2)
            y = (frmCustom.Top \ Screen.TwipsPerPixelY + (frmCustom.Height \ Screen.TwipsPerPixelY - (R.Bottom - R.Top)) \ 2)
            SetWindowPos hwnd, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE
            SetWindowLong hwnd, GWL_WNDPROC, lWndProc
    End Select
    Dlg_WndProc = CallWindowProc(lWndProc, hwnd, Msg, wParam, lParam)
End Function

