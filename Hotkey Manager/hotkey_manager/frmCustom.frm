VERSION 5.00
Object = "{88A12890-3854-11D4-B7E8-005004BC2C86}#3.0#0"; "AxAOLCmd.ocx"
Begin VB.Form frmCustom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Actions"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   Icon            =   "frmCustom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbHotkeys 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.ComboBox cmbActions 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   285
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin AxAOLCmd.AxAOLCmdCtl AxAOLCmdCtl1 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Cancel"
      Style           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCustom.frx":000C
      StandardColors  =   0   'False
      BorderDark      =   0
      BackColorClick  =   -2147483635
   End
   Begin AxAOLCmd.AxAOLCmdCtl AxAOLCmdCtl2 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "OK"
      Style           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCustom.frx":0028
      StandardColors  =   0   'False
      BorderDark      =   0
      BackColorClick  =   -2147483635
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hotkey:"
      Height          =   195
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Action:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents objTimer As clsTimer
Attribute objTimer.VB_VarHelpID = -1
Dim objFlatControl(1 To 3) As clsFlatControl

Private Const OFN_ALLOWMULTISELECT As Long = &H200
Private Const OFN_CREATEPROMPT As Long = &H2000
Private Const OFN_ENABLEHOOK As Long = &H20
Private Const OFN_ENABLETEMPLATE As Long = &H40
Private Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_EXTENSIONDIFFERENT As Long = &H400
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_LONGNAMES As Long = &H200000
Private Const OFN_NOCHANGEDIR As Long = &H8
Private Const OFN_NODEREFERENCELINKS As Long = &H100000
Private Const OFN_NOLONGNAMES As Long = &H40000
Private Const OFN_NONETWORKBUTTON As Long = &H20000
Private Const OFN_NOREADONLYRETURN As Long = &H8000
Private Const OFN_NOTESTFILECREATE As Long = &H10000
Private Const OFN_NOVALIDATE As Long = &H100
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_READONLY As Long = &H1
Private Const OFN_SHAREAWARE As Long = &H4000
Private Const OFN_SHAREFALLTHROUGH As Long = 2
Private Const OFN_SHAREWARN As Long = 0
Private Const OFN_SHARENOWARN As Long = 1
Private Const OFN_SHOWHELP As Long = &H10
Private Const OFS_MAXPATHNAME As Long = 260

'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not
'a standard Win32 type.
Private Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_CREATEPROMPT _
             Or OFN_NODEREFERENCELINKS

Private Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_OVERWRITEPROMPT _
             Or OFN_HIDEREADONLY

Private Type OPENFILENAME
  nStructSize       As Long
  hWndOwner         As Long
  hInstance         As Long
  sFilter           As String
  sCustomFilter     As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  sFile             As String
  nMaxFile          As Long
  sFileTitle        As String
  nMaxTitle         As Long
  sInitialDir       As String
  sDialogTitle      As String
  flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  sDefFileExt       As String
  nCustData         As Long
  fnHook            As Long
  sTemplateName     As String
End Type

Private OFN As OPENFILENAME

Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" _
   (pOpenfilename As OPENFILENAME) As Long

Private Sub AxAOLCmdCtl1_Click()
    Unload Me
End Sub

Private Sub AxAOLCmdCtl2_Click()
    'On Error GoTo LocalErr
    If txtFileName = "" Then
        MsgBox "Please don't let the file name empty.", vbInformation
        Exit Sub
    End If
    With frmHotkey.AxGridCtl2
        Dim strAction As String
        Dim Rs As Recordset
        Dim lMainID As Long
        Dim strHotkey As String
        strAction = cmbActions.Text & " " & txtFileName
        If .Tag = "" Then
            Set Rs = New Recordset
            Rs.Open "select_mainid_for_custom", Cn, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
            lMainID = Rs!mainid
        Else
            lMainID = CLng(.Tag)
        End If
        If cmbHotkeys.Text = "" Then
            strHotkey = "null"
        Else
            strHotkey = "'" & cmbHotkeys.Text & "'"
        End If
        Cn.Execute "exec insert_custom " & lMainID & ",'" & _
            strAction & "'," & strHotkey & ",'" & txtFileName & "'"
        Set Rs = New Recordset
        Rs.Open "last_id", Cn, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
        .AddRow
        .CellDetails .Rows, 1, cmbActions.Text & " " & txtFileName
        .CellDetails .Rows, 3, Rs(0)
        If cmbHotkeys.Text <> "" Then
            .CellDetails .Rows, 2, cmbHotkeys.Text
            frmHotkey.objHotKey.UnregisterKey CStr(Rs(0))
            frmHotkey.RegisterHotkey CStr(Rs(0)), cmbHotkeys.Text
        End If
        .AutoHeightRow .Rows, 18
    End With
    Unload Me
    Exit Sub
LocalErr:
    If Cn.Errors(0).Number = -2147467259 Then
        MsgBox "System found that you tried to assign a hotkey that has been already " & _
            "assigned to other action. You should assign another unused hotkey to the " & _
            "action", vbInformation, "Duplicate Hotkey"
    End If
End Sub

Private Sub Command1_Click()
    'used in call setup
    Dim sFilters As String
   
    'used after call
    Dim pos As Long
    Dim buff As String
    Dim sLongname As String
    Dim sShortname As String

    'create a string of filters for the dialog
    sFilters = "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    With OFN
        'size of the OFN structure
        .nStructSize = Len(OFN)
        'window owning the dialog
        .hWndOwner = hwnd
        'filters (patterns) for the dropdown combo
        .sFilter = sFilters
        'index to the initial filter
        'default filename, plus additional padding
        'for the user's final selection(s). Must be
        'double-null terminated
        .sFile = "" & Space$(1024) & vbNullChar & vbNullChar
        .nFilterIndex = 2
        'the size of the buffer
        .nMaxFile = Len(.sFile)
        'space for the file title if a single selection
        'made, double-null terminated, and its size
        .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
        .nMaxTitle = Len(OFN.sFileTitle)
        'the dialog title
        .sDialogTitle = "Select A File"
        'default open flags and multiselect
        .flags = OFS_FILE_OPEN_FLAGS Or OFN_FILEMUSTEXIST Or _
            OFN_HIDEREADONLY Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST
    End With
    'call the API
    hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
    If GetOpenFileName(OFN) Then txtFileName = OFN.sFile
    If hHook Then UnhookWindowsHookEx hHook
End Sub

Private Sub Form_Activate()
    Set objTimer = New clsTimer
    objTimer.Interval = 50
End Sub

Private Sub Form_Load()
    Dim i As Byte
    For i = 1 To 3
        Set objFlatControl(i) = New clsFlatControl
    Next
    objFlatControl(1).Attach txtFileName
    objFlatControl(2).Attach cmbActions
    objFlatControl(3).Attach cmbHotkeys
    cmbActions.AddItem "Open"
    cmbActions.AddItem "Edit"
    cmbActions.AddItem "Play"
    cmbActions.ListIndex = 0
    CopyComboToCombo frmHotkey.Combo1.hwnd, cmbHotkeys.hwnd
End Sub

Private Sub CopyComboToCombo(SourceHwnd As Long, TargetHwnd As Long)
    Dim c As Long
    Const CB_GETCOUNT = &H146
    Const CB_ADDSTRING = &H143
    Const CB_GETLBTEXT = &H148
    Dim numitems As Long
    Dim sItemText As String * 255
    LockWindowUpdate TargetHwnd
    numitems = SendMessageLong(SourceHwnd, CB_GETCOUNT, 0&, 0&)
    If numitems > 0 Then
        For c = 0 To numitems - 1
            Call SendMessageStr(SourceHwnd, CB_GETLBTEXT, c, ByVal sItemText)
            Call SendMessageStr(TargetHwnd, CB_ADDSTRING, 0&, ByVal sItemText)
        Next
    End If
    LockWindowUpdate 0&
End Sub

Private Sub Form_Unload(Cancel As Integer)
    objTimer.PulseTimer
    Set objTimer = Nothing
    Dim i As Byte
    For i = 1 To 3
        Set objFlatControl(i) = Nothing
    Next
End Sub

Private Sub objTimer_ThatTime()
    Dim CurPos As POINTAPI
    Dim lHwnd As Long
    Static lPreviousHwnd As Long
    Call GetCursorPos(CurPos)
    lHwnd = GetWindowFromPoint(hwnd, CurPos.X, CurPos.Y)
    If lPreviousHwnd = lHwnd Then Exit Sub
    Dim i As Byte
    Dim bIsTheButton As Boolean
    On Error Resume Next
    If AxAOLCmdCtl1.hwnd = lHwnd Then
        AxAOLCmdCtl1.BackColor = &HFF0000
        AxAOLCmdCtl2.BackColor = &HAA6D00
    ElseIf AxAOLCmdCtl2.hwnd = lHwnd Then
        AxAOLCmdCtl1.BackColor = &HAA6D00
        AxAOLCmdCtl2.BackColor = &HFF0000
    Else
        AxAOLCmdCtl1.BackColor = &HAA6D00
        AxAOLCmdCtl2.BackColor = &HAA6D00
    End If
    lPreviousHwnd = lHwnd
End Sub
