VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHitTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements clsISubclass

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Const WM_NCHITTEST = &H84&
Public Enum EHitTestAreas
   HTERROR = (-2)
   HTTRANSPARENT = (-1)
   HTNOWHERE = 0
   HTCLIENT = 1
   HTCAPTION = 2
   HTSYSMENU = 3
   HTGROWBOX = 4
   HTMENU = 5
   HTHSCROLL = 6
   HTVSCROLL = 7
   HTMINBUTTON = 8
   HTMAXBUTTON = 9
   HTLEFT = 10
   HTRIGHT = 11
   HTTOP = 12
   HTTOPLEFT = 13
   HTBOTTOM = 15
   HTBOTTOMLEFT = 16
   HTBOTTOMRIGHT = 17
   HTBORDER = 18
End Enum

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

' Array of objects to perform WM_NCHITTEST on:
Private m_lHTPtr() As Long
' Array of hWnds to perform WM_NCHITTEST on:
Private m_hWnd() As Long
' Number of objects to perform WM_NCHITTEST on:
Private m_iCount As Long

' Rectangles to perform hit testing on:
Private Type tHitTestRectangle
   hWnd As Long
   sKey As String
   lLeft As Long
   lTOp As Long
   lRight As Long
   lBottom As Long
   eResponse As EHitTestAreas
End Type
Private m_tHTR() As tHitTestRectangle
Private m_iHTRCount As Long

Public Sub SetTestArea( _
      ByVal vKey As Variant, _
      ByVal lLeftPixels As Long, _
      ByVal lTopPixels As Long, _
      ByVal lWidthPixels As Long, _
      ByVal lHeightPixels As Long, _
      ByVal eResponse As EHitTestAreas _
   )
Attribute SetTestArea.VB_Description = "Sets the rectangle to be hit tested and response for an area added earlier with the AddArea method."
Dim iItem As Long

   iItem = AreaIndex(vKey)
   If (iItem > 0) Then
      With m_tHTR(iItem)
         .lLeft = lLeftPixels
         .lTOp = lTopPixels
         .lRight = lLeftPixels + lWidthPixels
         .lBottom = lTopPixels + lHeightPixels
         .eResponse = eResponse
      End With
   End If
   
End Sub
Public Sub SetTestAreaFromObject( _
      ByVal vKey As Variant, _
      ByRef objThis As Object, _
      ByVal eResponse As EHitTestAreas _
   )
Attribute SetTestAreaFromObject.VB_Description = "Similar to the SetTestArea method but sets the rectangle to be hit tested directly from an object such as Label control."
   SetTestArea vKey, objThis.Left \ Screen.TwipsPerPixelX, objThis.Top \ Screen.TwipsPerPixelY, objThis.Width \ Screen.TwipsPerPixelY, objThis.Height \ Screen.TwipsPerPixelY, eResponse
End Sub

Public Function AddArea( _
      ByVal hWnd As Long, _
      Optional ByVal sKey As String = "" _
   ) As Long
Attribute AddArea.VB_Description = "Adds an area to be hit tested."
      
   m_iHTRCount = m_iHTRCount + 1
   ReDim Preserve m_tHTR(1 To m_iHTRCount) As tHitTestRectangle
   AddArea = m_iHTRCount
   With m_tHTR(m_iHTRCount)
      .hWnd = hWnd
      If (sKey <> "") Then
         .sKey = sKey
      Else
         .sKey = CStr(m_iHTRCount)
      End If
   End With

End Function

Private Function AreaIndex(ByVal vKey As Variant) As Long
Dim i As Long

   If (IsNumeric(vKey)) Then
      i = CLng(vKey)
      If (i > 0) And (i <= m_iHTRCount) Then
         AreaIndex = i
         Exit Function
      End If
   Else
      For i = 1 To m_iHTRCount
         If (m_tHTR(i).sKey = vKey) Then
            AreaIndex = i
            Exit Function
         End If
      Next i
   End If
   
   ' error
End Function

Public Sub AttachHitTester(ByRef objThis As Object)
Attribute AttachHitTester.VB_Description = "Adds a window to the list of windows to perform hit testing on."
Dim hWnd As Long
   hWnd = objThis.hWnd
   If (IsNew(hWnd)) Then
      m_iCount = m_iCount + 1
      ReDim Preserve m_lHTPtr(1 To m_iCount) As Long
      ReDim Preserve m_hWnd(1 To m_iCount) As Long
      m_lHTPtr(m_iCount) = ObjPtr(objThis)
      m_hWnd(m_iCount) = hWnd
      AttachMessage Me, m_hWnd(m_iCount), WM_NCHITTEST
   End If
End Sub

Public Sub DestroyHitTester()
Attribute DestroyHitTester.VB_Description = "Clears up all hit testers.  Called automatically when the object is destroyed."
Dim i As Long
   For i = 1 To m_iCount
      If (m_hWnd(i) <> 0) Then
         DetachMessage Me, m_hWnd(i), WM_NCHITTEST
         m_hWnd(i) = 0
      End If
      m_lHTPtr(i) = 0
   Next i
   m_iCount = 0
End Sub

Private Function IsNew(ByVal hWnd As Long) As Boolean
Dim i As Long
   IsNew = True
   For i = 1 To m_iCount
      If (m_hWnd(i) = hWnd) Then
         IsNew = False
         Exit For
      End If
   Next i
      
End Function

Public Property Get PointInObject(ByRef oThis As Object, ByVal xPixels As Long, ByVal yPixels As Long) As Boolean
Attribute PointInObject.VB_Description = "Helper function.  Returns whether the given point is within the boundary of an object."
Dim tR As RECT
Dim tFR As RECT

   On Error Resume Next
   ' Get object position:
   tR.Left = oThis.Left \ Screen.TwipsPerPixelX
   tR.Right = tR.Left + oThis.Width \ Screen.TwipsPerPixelX
   tR.Top = oThis.Top \ Screen.TwipsPerPixelX
   tR.Bottom = tR.Top + oThis.Height \ Screen.TwipsPerPixelY
   If (Err.Number = 0) Then
      ' Check for hit test:
      PointInObject = (PtInRect(tR, xPixels, yPixels) <> 0)
   End If
End Property

Private Property Let clsISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get clsISubclass_MsgResponse() As EMsgResponse
   clsISubclass_MsgResponse = emrPreprocess
End Property

Private Function clsISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tP As POINTAPI
Dim tR As RECT
Dim lR As Long
Dim bDoDefault As Boolean
Dim eType As EHitTestAreas
Dim i As Long, j As Long

   For j = 1 To m_iHTRCount
      With m_tHTR(j)
         If (.hWnd = hWnd) Then
            GetCursorPos tP
            ScreenToClient hWnd, tP
            eType = HTCLIENT
            tR.Left = .lLeft
            tR.Top = .lTOp
            tR.Right = .lRight
            tR.Bottom = .lBottom
            If (PtInRect(tR, tP.x, tP.y) <> 0) Then
               clsISubclass_WindowProc = .eResponse
               Exit Function
            End If
         End If
      End With
   Next j
   
   ' Do the default:
   clsISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
   
End Function

