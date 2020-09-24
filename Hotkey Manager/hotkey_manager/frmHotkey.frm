VERSION 5.00
Object = "{88A12890-3854-11D4-B7E8-005004BC2C86}#3.0#0"; "AxAOLCmd.ocx"
Object = "{4F387820-387D-11D4-B7E9-005004BC2C86}#3.0#0"; "AxGrid.ocx"
Object = "{88A12862-3854-11D4-B7E8-005004BC2C86}#3.0#0"; "AxImageList.ocx"
Object = "{88A128A4-3854-11D4-B7E8-005004BC2C86}#3.0#0"; "AxTray.ocx"
Begin VB.Form frmHotkey 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8520
   Icon            =   "frmHotkey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin AxTray.AxTrayCtl AxTrayCtl1 
      Left            =   1800
      Top             =   2040
      _ExtentX        =   900
      _ExtentY        =   900
   End
   Begin AxAOLCmd.AxAOLCmdCtl AxAOLCmdCtl1 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   0
      Left            =   6960
      TabIndex        =   0
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Stop"
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
      Picture         =   "frmHotkey.frx":0442
      StandardColors  =   0   'False
      BorderDark      =   0
      BackColorClick  =   -2147483635
   End
   Begin AxImageList.AxImageListCtl AxImageListCtl1 
      Left            =   3600
      Top             =   1800
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   1880
      Images          =   "frmHotkey.frx":045E
      KeyCount        =   2
      Keys            =   "ÿ"
   End
   Begin AxAOLCmd.AxAOLCmdCtl AxAOLCmdCtl1 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   1
      Left            =   5520
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Send to Tray"
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
      Picture         =   "frmHotkey.frx":0BD6
      StandardColors  =   0   'False
      BorderDark      =   0
      BackColorClick  =   -2147483635
   End
   Begin AxAOLCmd.AxAOLCmdCtl AxAOLCmdCtl1 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   2
      Left            =   4080
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "About"
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
      Picture         =   "frmHotkey.frx":0BF2
      StandardColors  =   0   'False
      BorderDark      =   0
      BackColorClick  =   -2147483635
   End
   Begin AxAOLCmd.AxAOLCmdCtl AxAOLCmdCtl1 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   6
      Left            =   1680
      TabIndex        =   5
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Custom Actions"
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
      Picture         =   "frmHotkey.frx":0C0E
      StandardColors  =   0   'False
      BorderDark      =   0
      BackColorClick  =   -2147483635
   End
   Begin AxAOLCmd.AxAOLCmdCtl AxAOLCmdCtl1 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   7
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Standard Actions"
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
      Picture         =   "frmHotkey.frx":0C2A
      BackColor       =   8388736
      StandardColors  =   0   'False
      BorderDark      =   0
      BackColorClick  =   -2147483635
   End
   Begin AxAOLCmd.AxAOLCmdCtl AxAOLCmdCtl1 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   5
      Left            =   5520
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Add Item"
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
      Picture         =   "frmHotkey.frx":0C46
      StandardColors  =   0   'False
      BorderDark      =   0
      BackColorClick  =   -2147483635
   End
   Begin AxAOLCmd.AxAOLCmdCtl AxAOLCmdCtl1 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   3
      Left            =   6960
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Delete Item"
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
      Picture         =   "frmHotkey.frx":0C62
      StandardColors  =   0   'False
      BorderDark      =   0
      BackColorClick  =   -2147483635
   End
   Begin AxGrid.AxGridCtl AxGridCtl2 
      Height          =   3015
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   8050
      _ExtentX        =   14208
      _ExtentY        =   5318
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisableIcons    =   -1  'True
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin AxGrid.AxGridCtl AxGridCtl1 
      Height          =   3615
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   8050
      _ExtentX        =   14208
      _ExtentY        =   6376
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisableIcons    =   -1  'True
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmHotkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objFlatControl(1 To 2) As clsFlatControl
Dim WithEvents objTimer As clsTimer
Attribute objTimer.VB_VarHelpID = -1

Public WithEvents objHotKey As clsRegHotKey
Attribute objHotKey.VB_VarHelpID = -1

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Enum ShowType
    SW_SHOWNORMAL = 1
    SW_SHOWMAXIMIZED = 3
    SW_SHOWDEFAULT = 10
End Enum

Private objHitTest As New clsHitTester

Private Sub Buildgrid2()
    Dim Rs As Recordset
    Set Rs = New Recordset
    Rs.Open "custom", Cn, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    Set Rs.ActiveConnection = Nothing
    With AxGridCtl2
        ' Turn redraw off for speed
        .Redraw = False
        .Editable = True
        .ImageList = AxImageListCtl1
        ' Row mode - select the entire row
        .RowMode = True
        .GridLines = True
        .GridLineColor = vbYellow
        ' Outlook style for the header control
        .HeaderFlat = True
        If Not Rs.EOF Then .Tag = CStr(Rs!mainid)
        .AddColumn vkey:="desc", sHeader:="Description/Action", lcolumnwidth:=300, ealign:=ecgHdrTextALignCentre
        .AddColumn vkey:="key", sHeader:="Shortcut Key", lcolumnwidth:=200, ealign:=ecgHdrTextALignCentre
        .AddColumn vkey:="id", lcolumnwidth:=16, bvisible:=False, bincludeinselect:=False
        Do Until Rs.EOF
            .AddRow
            .CellDetails .Rows, 1, Rs!actions
            .CellDetails .Rows, 3, Rs!id
            If Not IsNull(Rs!hotkey) Then
                .CellDetails .Rows, 2, Rs!hotkey
                objHotKey.UnregisterKey CStr(Rs!id)
                RegisterHotkey CStr(Rs!id), Rs!hotkey
            End If
            .AutoHeightRow .Rows, 18
            Rs.MoveNext
        Loop
        Set .BackgroundPicture = LoadPictureResource(102, "JPEG", False)
        .Redraw = True
    End With
End Sub

Private Sub BuildGrid1()
    Dim strCurrentOS As String
    Dim Rs As Recordset
    Set Cn = New Connection
    Cn.CursorLocation = adUseClient
    Cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\hotkeys.mdb;Persist Security Info=False"
    strCurrentOS = OS
    Set Rs = New Recordset
    Rs.Open strCurrentOS, Cn, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    Set Rs.ActiveConnection = Nothing
    With AxGridCtl1
        ' Turn redraw off for speed
        .Redraw = False
        .Editable = True
        .ImageList = AxImageListCtl1
        ' Row mode - select the entire row
        .RowMode = True
        .GridLines = True
        .GridLineColor = vbYellow
        ' Outlook style for the header control
        .HeaderFlat = True
        ' Add the columns
        .AddColumn vkey:="group1", lcolumnwidth:=16, bvisible:=False, bincludeinselect:=False
        .AddColumn vkey:="desc", sHeader:="Description/Action", lcolumnwidth:=400, ealign:=ecgHdrTextALignCentre
        .AddColumn vkey:="key", sHeader:="Shortcut Key", lcolumnwidth:=200, ealign:=ecgHdrTextALignCentre
        .AddColumn vkey:="id", lcolumnwidth:=16, bvisible:=False, bincludeinselect:=False
        .AddColumn vkey:="main", lcolumnwidth:=96 + 256 + 96 + 96, bvisible:=False, bincludeinselect:=False, browtextcolumn:=True
        .KeySearchColumn = .ColumnIndex("desc")
        .SetHeaders
        ' Set up a bold font
        Dim sFnt As New StdFont
        sFnt.Name = "Arial"
        sFnt.Size = 8
        Do Until Rs.EOF
            .AddRow
            .CellDetails .Rows, 4, Rs!id
            .CellDetails .Rows, 2, Rs!actions, , , , , sFnt
            .CellDetails .Rows, 5, Rs!main
            If Not IsNull(Rs!hotkey) Then
                .CellDetails .Rows, 3, Rs!hotkey
                objHotKey.UnregisterKey CStr(Rs!id)
                RegisterHotkey CStr(Rs!id), Rs!hotkey
            End If
            .AutoHeightRow .Rows, 18
            Rs.MoveNext
        Loop
        .AutoWidthColumn "desc"
        Dim sThis(0) As String
        Dim eOrder(0) As cShellSortOrderCOnstants
        sThis(0) = "main"
        eOrder(0) = CCLOrderAscending
        DoGroup 1, sThis(), eOrder()
        Set .BackgroundPicture = LoadPictureResource(102, "JPEG", False)
        .Redraw = True
        .ZOrder 0
        .Visible = True
    End With
End Sub

Public Sub RegisterHotkey(strName As String, strHotkey As String)
    Dim Modifiers As EHKModifiers
    If InStr(1, strHotkey, "CTRL") Then
        Modifiers = MOD_CONTROL
    End If
    If InStr(1, strHotkey, "SHIFT") Then
        Modifiers = Modifiers + MOD_SHIFT
    End If
    If InStr(1, strHotkey, "ALT") Then
        Modifiers = Modifiers + MOD_ALT
    End If
    Dim strLastTwoChar As String
    Dim iKeyCode As KeyCodeConstants
    strLastTwoChar = Right(strHotkey, 2)
    If Mid(strLastTwoChar, 1, 1) = "+" Then
        objHotKey.RegisterKey strName, Asc(Mid(strLastTwoChar, 2, 1)), Modifiers
    Else
        Select Case strLastTwoChar
            Case Is = "F1"
                iKeyCode = vbKeyF1
            Case Is = "F2"
                iKeyCode = vbKeyF2
            Case Is = "F3"
                iKeyCode = vbKeyF3
            Case Is = "F4"
                iKeyCode = vbKeyF4
            Case Is = "F5"
                iKeyCode = vbKeyF5
            Case Is = "F6"
                iKeyCode = vbKeyF6
            Case Is = "F7"
                iKeyCode = vbKeyF7
            Case Is = "F8"
                iKeyCode = vbKeyF8
            Case Is = "F9"
                iKeyCode = vbKeyF9
            Case Is = "10"
                iKeyCode = vbKeyF10
            Case Is = "11"
                iKeyCode = vbKeyF11
            Case Is = "12"
                iKeyCode = vbKeyF12
        End Select
        objHotKey.RegisterKey strName, iKeyCode, Modifiers
    End If
End Sub

Private Sub AxAOLCmdCtl1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            Visible = False
            AxTrayCtl1.Visible = True
        Case 2
            Load frmAbout
            DoEvents
            frmAbout.Show vbModal
        Case 3
            With AxGridCtl2
                If .SelectedRow > 0 Then
                    If MsgBox("Are you sure you want to delete the selected item?") Then
                        Cn.Execute "exec delete_custom " & .CellText(.SelectedRow, 3)
                        objHotKey.UnregisterKey .CellText(.SelectedRow, 3)
                        .RemoveRow .SelectedRow
                    End If
                End If
            End With
        Case 5
            Load frmCustom
            DoEvents
            frmCustom.Show vbModal
        Case 6
            AxGridCtl1.Visible = False
            MakeVisible True
            AxAOLCmdCtl1(6).BackColor = &H800080
            AxAOLCmdCtl1(7).BackColor = &HAA6D00
            AxGridCtl2.Visible = True
        Case 7
            AxGridCtl2.Visible = False
            MakeVisible False
            AxAOLCmdCtl1(6).BackColor = &HAA6D00
            AxAOLCmdCtl1(7).BackColor = &H800080
            AxGridCtl1.Visible = True
    End Select
End Sub

Private Sub MakeVisible(bVal As Boolean)
    AxAOLCmdCtl1(3).Visible = bVal
    AxAOLCmdCtl1(5).Visible = bVal
End Sub

Private Sub AxGridCtl1_CancelEdit()
    Combo1.Visible = False
End Sub

Private Sub AxGridCtl1_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    Dim sKey As String
    Dim bFound As Boolean
    Dim lItemData As Long
    Dim bIgnoreUntilNext As Boolean
    With AxGridCtl1
        If (lRow > 0) And (lCol > 0) Then
           ' Dbl clicked on a valid cell.  Find out whether it is a group or
           ' not:
           sKey = .ColumnKey(lCol)
           If (sKey = "main") Then
              .Redraw = False
              ' Expand or collapse:
              lItemData = .CellItemData(lRow, 5)
              If (.CellExtraIcon(lRow, 5) = 0) Then
                 ' collapse:
                 .CellExtraIcon(lRow, 5) = 1
                 lRow = lRow + 1
                 Do While lRow <= .Rows And Not bFound
                    If .CellItemData(lRow, 5) = 0 Or .CellItemData(lRow, 5) > lItemData Then
                       .RowVisible(lRow) = False
                    Else
                       bFound = True
                    End If
                    lRow = lRow + 1
                 Loop
              Else
                 ' expand:
                 .CellExtraIcon(lRow, 5) = 0
                 lRow = lRow + 1
                 Do While lRow <= .Rows And Not bFound
                    If .CellItemData(lRow, 5) = 0 Then
                       If Not (bIgnoreUntilNext) Then
                          .RowVisible(lRow) = True
                       End If
                    ElseIf .CellItemData(lRow, 5) > lItemData Then
                       .RowVisible(lRow) = True
                       bIgnoreUntilNext = (.CellExtraIcon(lRow, 5) = 1)
                    Else
                       bFound = True
                    End If
                    lRow = lRow + 1
                 Loop
              End If
              .Redraw = True
           End If
        End If
    End With
End Sub

Private Sub AxGridCtl1_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
    Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
    Dim iArt As Long, iRow As Long, iType As Long, iArticle As Long, iLink As Long
    Dim bDontAdd As Boolean
    If lCol = 3 Then
        AxGridCtl1.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
        With Combo1
            .Move lLeft, lTop, lWidth
            If AxGridCtl1.CellText(lRow, 3) <> "" Then
                .Text = AxGridCtl1.CellText(lRow, 3)
            Else
                .ListIndex = -1
            End If
            .Visible = True
            .ZOrder
            .SetFocus
        End With
    Else
        bCancel = True
    End If
End Sub

Private Sub AxGridCtl2_CancelEdit()
    Combo2.Visible = False
End Sub

Private Sub AxGridCtl2_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
    Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
    Dim iArt As Long, iRow As Long, iType As Long, iArticle As Long, iLink As Long
    Dim bDontAdd As Boolean
    If lCol = 2 Then
        AxGridCtl2.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
        With Combo2
            .Move lLeft, lTop, lWidth
            If AxGridCtl2.CellText(lRow, 2) <> "" Then
                .Text = AxGridCtl2.CellText(lRow, 2)
            Else
                .ListIndex = -1
            End If
            .Visible = True
            .ZOrder
            .SetFocus
        End With
    Else
        bCancel = True
    End If
End Sub

Private Sub AxTrayCtl1_DblClick()
    AxTrayCtl1.Visible = False
    Visible = True
End Sub

Private Sub Combo1_Click()
    If Combo1.ListIndex > -1 And Combo1.Visible Then
        With AxGridCtl1
            If Combo1.Text <> "" Then
                On Error GoTo LocalErr
                Cn.Execute "exec insert_hotkey '" & Combo1.Text & "','" & AxGridCtl1.CellText(AxGridCtl1.SelectedRow, 4) & "'"
            Else
                On Error GoTo LocalErr
                Cn.Execute "exec insert_hotkey null,'" & AxGridCtl1.CellText(AxGridCtl1.SelectedRow, 4) & "'"
            End If
            .CellText(.SelectedRow, 3) = Combo1.Text
            objHotKey.UnregisterKey AxGridCtl1.CellText(AxGridCtl1.SelectedRow, 4)
            RegisterHotkey AxGridCtl1.CellText(AxGridCtl1.SelectedRow, 4), Combo1.Text
            .SetFocus
        End With
    End If
    Exit Sub
LocalErr:
    If Cn.Errors(0).Number = -2147467259 Then
        MsgBox "System found that you tried to assign a hotkey that has been already " & _
            "assigned to other action. You should assign another unused hotkey to the " & _
            "action", vbInformation, "Duplicate Hotkey"
    End If
    AxGridCtl1.CancelEdit
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyEscape) Then
        Combo1.ListIndex = -1
        AxGridCtl1.SetFocus
    End If
End Sub

Private Sub Combo1_LostFocus()
    AxGridCtl1.CancelEdit
End Sub

Private Sub Combo2_Click()
    If Combo2.ListIndex > -1 And Combo2.Visible Then
        With AxGridCtl2
            If Combo2.Text <> "" Then
                On Error GoTo LocalErr
                Cn.Execute "exec insert_hotkey '" & Combo2.Text & "','" & .CellText(.SelectedRow, 3) & "'"
            End If
            .CellText(.SelectedRow, 2) = Combo2.Text
            objHotKey.UnregisterKey .CellText(.SelectedRow, 3)
            RegisterHotkey .CellText(.SelectedRow, 3), Combo2.Text
            .SetFocus
        End With
    End If
    Exit Sub
LocalErr:
    If Cn.Errors(0).Number = -2147467259 Then
        MsgBox "System found that you tried to assign a hotkey that has been already " & _
            "assigned to other action. You should assign another unused hotkey to the " & _
            "action", vbInformation, "Duplicate Hotkey"
    End If
    AxGridCtl2.CancelEdit
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyEscape) Then
        Combo2.ListIndex = -1
        AxGridCtl2.SetFocus
    End If
End Sub

Private Sub Combo2_LostFocus()
    AxGridCtl2.CancelEdit
End Sub

Private Sub Form_Activate()
    Set objTimer = New clsTimer
    objTimer.Interval = 50
End Sub

Private Sub Form_Deactivate()
    objTimer.PulseTimer
End Sub

Private Sub Form_Load()
    Set AxTrayCtl1.Icon = Me.Icon
    AxTrayCtl1.Visible = True
    AxTrayCtl1.ToolTip = "Hotkey Manager Beta Version"
    Picture = LoadPictureResource(101, "JPEG", False)
    Set objHotKey = New clsRegHotKey
    objHotKey.Attach hwnd
    BuildGrid1
    Buildgrid2
    LoadCombo
    Set objFlatControl(1) = New clsFlatControl
    Set objFlatControl(2) = New clsFlatControl
    objFlatControl(1).Attach Combo1
    objFlatControl(2).Attach Combo2
    With objHitTest
        .AttachHitTester Me
        .AddArea hwnd, "me"
        .SetTestAreaFromObject "me", Me, HTCAPTION
    End With
End Sub

Private Sub LoadCombo()
    Dim Rs As Recordset
    Const CB_ADDSTRING = &H143
    Set Rs = New Recordset
    Rs.Open "all_available_keys", Cn, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    Combo1.Clear
    Combo2.Clear
    LockWindowUpdate Combo1.hwnd
    LockWindowUpdate Combo2.hwnd
    Call SendMessageStr(Combo1.hwnd, CB_ADDSTRING, 0&, ByVal "")
    Call SendMessageStr(Combo2.hwnd, CB_ADDSTRING, 0&, ByVal "")
    Do Until Rs.EOF
        Call SendMessageStr(Combo1.hwnd, CB_ADDSTRING, 0&, ByVal Rs!available_hotkeys)
        Call SendMessageStr(Combo2.hwnd, CB_ADDSTRING, 0&, ByVal Rs!available_hotkeys)
        'Combo1.AddItem Rs!available_hotkeys
        Rs.MoveNext
    Loop
    LockWindowUpdate 0&
End Sub

Private Function OS() As String
    Dim lngBuild As Long, pos As Integer
    Dim strVersion As String
    Dim objOSVerInfo As clsOSVerInfo
    Set objOSVerInfo = New clsOSVerInfo
    objOSVerInfo.WindowsVersion
    With objOSVerInfo
        Select Case .dwPlatformId
            Case VER_PLATFORM_WIN32_WINDOWS
                If .dwMajorVersion = 4 Then
                    If .dwMinorVersion <> 0 Then
                        OS = "Win98"
                    Else
                        OS = "Win95"
                    End If
                ElseIf .dwMajorVersion = 5 Then
                    OS = "Win2000"
                End If
            Case ver_platform_win32_nt
                OS = "WinNT"
        End Select
    End With
    Set objOSVerInfo = Nothing
End Function

Public Sub DoGroup(ByVal iItems As Long, sGroupColumns() As String, eOrder() As cShellSortOrderCOnstants)
    Dim i As Long
    Dim iRow As Long
    Dim iCol As Long
    Dim iNumber As Long
    Dim sFnt As StdFont
    Dim iFnt As IFont
    Dim sJunk() As String, eJunk() As cShellSortOrderCOnstants
    Dim bForce As Boolean
    Static iRefCount As Long
    With AxGridCtl1
        iRefCount = iRefCount + 1
        iNumber = iItems - 1
        If (iNumber > 2) Then
            MsgBox "Can't do it - max grouping is restricted to 3 columns for this demo.", vbInformation
        Else
            ' Stop redraw for speed:
            If (iRefCount = 1) Then
                .Redraw = False
            End If
            If (iNumber < 0) Then
                ' Remove all existing group rows:
                For iRow = .Rows To 1 Step -1
                    If (.CellItemData(iRow, 5) > 0) Then
                        .RemoveRow iRow
                    End If
                Next iRow
                For iRow = 1 To .Rows
                    .RowVisible(iRow) = True
                Next iRow
            Else
            ' Remove groupings:
            DoGroup 0, sJunk(), eJunk()
            ' Make the relevant headers visible:
            If (0 <= iNumber) Then
                .ColumnVisible("group1") = True
            End If
            ' Sort the grid according to the groupings:
            With .SortObject
                .Clear
                For i = 0 To iNumber
                    .SortColumn(i + 1) = AxGridCtl1.ColumnIndex(sGroupColumns(i))
                    .SortOrder(i + 1) = eOrder(i)
                    .SortType(i + 1) = AxGridCtl1.ColumnSortType(sGroupColumns(i))
                Next i
            End With
            .Sort
            ' Now add grouping rows:
            ReDim vLastItem(0 To iNumber) As Variant
            Set iFnt = .Font
            iFnt.Clone sFnt
            sFnt.Bold = True
            iRow = 1
            Do
                bForce = False
                For i = 0 To iNumber
                    If Not .RowIsGroup(iRow) Then
                        iCol = .ColumnIndex(sGroupColumns(i))
                        If .CellText(iRow, iCol) <> vLastItem(i) Or bForce Then
                            vLastItem(i) = .CellText(iRow, iCol)
                            .AddRow iRow, "GROUP", , , True, i + 1
                            .CellDetails iRow, 5, vLastItem(i), , , vbButtonFace, , sFnt, , 1, i + 1
                            bForce = True
                        End If
                    End If
                Next i
                iRow = iRow + 1
            Loop While iRow < .Rows
            For iRow = 1 To .Rows
                If Not .CellItemData(iRow, 5) = 1 Then
                    .RowVisible(iRow) = False
                End If
            Next iRow
            End If
            ' Start redrawing again:
            If (iRefCount = 1) Then
                .Redraw = True
            End If
        End If
        iRefCount = iRefCount - 1
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    objTimer.PulseTimer
    Set objTimer = Nothing
    Cn.Close
    Set Cn = Nothing
    objHotKey.Clear
    Set objHotKey = Nothing
    Set objFlatControl(1) = Nothing
    Set objFlatControl(2) = Nothing
    objHitTest.DestroyHitTester
    Set objHitTest = Nothing
End Sub

Private Sub objHotKey_HotKeyPress(ByVal sName As String, ByVal eModifiers As AxHotkey.EHKModifiers, ByVal eKey As KeyCodeConstants)
    Dim Rs As New Recordset
    Dim strSQL As String
    strSQL = "exec select_hotkey '" & sName & "'"
    Rs.Open strSQL, Cn, adOpenForwardOnly, adCmdStoredProc
    If Not Rs.EOF Then
        Select Case Mid(Rs!actions, 1, 4)
            Case "Open", "Edit", "Play"
                RunShellExecute Mid(Rs!actions, 1, 4), Rs!Commands, 0&, 0&, SW_SHOWNORMAL
            Case Else
                Shell Rs!Commands, vbNormalFocus
        End Select
    End If
End Sub

Private Sub RunShellExecute(sTopic As String, sFile As Variant, sParams As Variant, sDirectory As Variant, nShowCmd As ShowType)
    'Shell default proSram to edit report
    'This depends on type of report
    Dim hWndDesk As Long
    Dim success As Long
    hWndDesk = GetDesktopWindow()
    success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)
    If success < 32 Then
       Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
    End If
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
    For i = 0 To 7
        If AxAOLCmdCtl1(i).hwnd = lHwnd Then
            bIsTheButton = True
            Exit For
        End If
    Next
    If bIsTheButton Then
        For i = 0 To 7
            If AxAOLCmdCtl1(i).hwnd = lHwnd Then
                AxAOLCmdCtl1(i).BackColor = &HFF0000
            Else
                Select Case i
                    Case 7
                        If AxGridCtl1.Visible Then
                            AxAOLCmdCtl1(i).BackColor = &H800080
                        Else
                            AxAOLCmdCtl1(i).BackColor = &HAA6D00
                        End If
                    Case 6
                        If AxGridCtl2.Visible Then
                            AxAOLCmdCtl1(i).BackColor = &H800080
                        Else
                            AxAOLCmdCtl1(i).BackColor = &HAA6D00
                        End If
                    Case Else
                        AxAOLCmdCtl1(i).BackColor = &HAA6D00
                End Select
            End If
        Next
    End If
    lPreviousHwnd = lHwnd
End Sub

