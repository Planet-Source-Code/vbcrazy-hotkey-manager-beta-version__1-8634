VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFlash.ocx"
Object = "{88A12890-3854-11D4-B7E8-005004BC2C86}#3.0#0"; "AxAOLCmd.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   ClientHeight    =   5145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   340.043
   ScaleMode       =   0  'User
   ScaleWidth      =   309.098
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin AxAOLCmd.AxAOLCmdCtl AxAOLCmdCtl2 
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "Who Am I?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmAbout.frx":0000
      StandardColors  =   0   'False
   End
   Begin AxAOLCmd.AxAOLCmdCtl AxAOLCmdCtl1 
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "Exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmAbout.frx":001C
      StandardColors  =   0   'False
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _cx             =   4207454
      _cy             =   4203432
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   0   'False
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   -1  'True
      BGColor         =   ""
      SWRemote        =   ""
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents objTimer As clsTimer
Attribute objTimer.VB_VarHelpID = -1

Private Sub AxAOLCmdCtl1_Click()
    Unload Me
End Sub

Private Sub AxAOLCmdCtl2_Click()
    If AxAOLCmdCtl2.Caption = "Who Am I?" Then
        LoadPictureResource 104, "SWF", True
        AxAOLCmdCtl2.Caption = "About"
    Else
        LoadPictureResource 103, "SWF", True
        AxAOLCmdCtl2.Caption = "Who Am I?"
    End If
End Sub

Private Sub Form_Activate()
    Set objTimer = New clsTimer
    objTimer.Interval = 50
End Sub

Private Sub Form_Load()
    LoadPictureResource 103, "SWF", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    objTimer.PulseTimer
    Set objTimer = Nothing
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

