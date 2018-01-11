VERSION 5.00
Object = "{DF6596FB-AFA4-489C-94B0-FAC1615747AA}#1.2#0"; "PagerCtlU.ocx"
Object = "{C481749B-3055-4751-A915-E5DE328C878E}#1.2#0"; "PagerCtlA.ocx"
Begin VB.Form frmMain 
   Caption         =   "Pager 1.2 - Events Sample"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   2  'Bildschirmmitte
   Begin PagerCtlLibACtl.Pager PagerA 
      Height          =   2895
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   7215
      _cx             =   12726
      _cy             =   5106
      Appearance      =   0
      AutoScrollFrequency=   -1
      AutoScrollLinesPerTimeout=   0
      AutoScrollPixelsPerLine=   0
      BackColor       =   -2147483633
      BorderSize      =   0
      BorderStyle     =   0
      DetectDoubleClicks=   -1  'True
      DisabledEvents  =   0
      Enabled         =   -1  'True
      ForwardMouseMessagesToContainedWindow=   0   'False
      HoverTime       =   -1
      MousePointer    =   0
      Orientation     =   1
      RegisterForOLEDragDrop=   0   'False
      ScrollAutomatically=   2
      ScrollButtonSize=   12
      SupportOLEDragImages=   -1  'True
      Begin VB.PictureBox Picture1 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'Kein
         Height          =   2055
         Index           =   1
         Left            =   840
         ScaleHeight     =   2055
         ScaleWidth      =   3255
         TabIndex        =   11
         Top             =   480
         Width           =   3255
         Begin VB.CommandButton Command1 
            Caption         =   "Command6"
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   18
            Top             =   3840
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command2"
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command3"
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   14
            Top             =   1680
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command4"
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   13
            Top             =   2400
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command5"
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   12
            Top             =   3120
            Width           =   1575
         End
      End
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "&Log"
      Height          =   255
      Left            =   7320
      TabIndex        =   2
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8430
      TabIndex        =   3
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox txtLog 
      Height          =   4815
      Left            =   7320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin PagerCtlLibUCtl.Pager PagerU 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _cx             =   12726
      _cy             =   4895
      Appearance      =   0
      AutoScrollFrequency=   -1
      AutoScrollLinesPerTimeout=   0
      AutoScrollPixelsPerLine=   0
      BackColor       =   -2147483633
      BorderSize      =   0
      BorderStyle     =   0
      DetectDoubleClicks=   -1  'True
      DisabledEvents  =   0
      Enabled         =   -1  'True
      ForwardMouseMessagesToContainedWindow=   0   'False
      HoverTime       =   -1
      MousePointer    =   0
      Orientation     =   1
      RegisterForOLEDragDrop=   0   'False
      ScrollAutomatically=   2
      ScrollButtonSize=   12
      SupportOLEDragImages=   -1  'True
      Begin VB.PictureBox Picture1 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'Kein
         Height          =   2055
         Index           =   0
         Left            =   840
         ScaleHeight     =   2055
         ScaleWidth      =   3255
         TabIndex        =   5
         Top             =   240
         Width           =   3255
         Begin VB.CommandButton Command1 
            Caption         =   "Command6"
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   17
            Top             =   3840
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command5"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   10
            Top             =   3120
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command4"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   9
            Top             =   2400
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command3"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   1680
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command2"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private bLog As Boolean
  Private objActiveCtl As Object


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidFrmMain
      lRet = HandleMessage_Form(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "frmMain.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_Form(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const WM_NOTIFYFORMAT = &H55
  Const WM_USER = &H400
  Const OCM__BASE = WM_USER + &H1C00
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_NOTIFYFORMAT
      ' give the control a chance to request Unicode notifications
      lRet = SendMessageAsLong(wParam, OCM__BASE + uMsg, wParam, lParam)

      bCallDefProc = False
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub chkLog_Click()
  bLog = (chkLog.Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  objActiveCtl.About
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  chkLog.Value = CheckBoxConstants.vbChecked

  ' this is required to make the control work as expected
  Subclass
  
  PagerU.hContainedWindow = Picture1(0).hWnd
  PagerA.hContainedWindow = Picture1(1).hWnd
End Sub

Private Sub Form_Resize()
  Dim cx As Long

  If Me.WindowState <> vbMinimized Then
    cx = 0.4 * Me.ScaleWidth
    txtLog.Move Me.ScaleWidth - cx, 0, cx, Me.ScaleHeight - cmdAbout.Height - 10
    cmdAbout.Move txtLog.Left + (cx / 2) - cmdAbout.Width / 2, txtLog.Top + txtLog.Height + 5
    chkLog.Move txtLog.Left, cmdAbout.Top + 5
    PagerU.Move 0, 10, txtLog.Left - 5, (Me.ScaleHeight - 20) / 2 - 90
    PagerA.Move 0, PagerU.Top + PagerU.Height + 20, PagerU.Width, PagerU.Height
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub PagerA_BeforeScroll(ByVal scrollDirection As PagerCtlLibACtl.ScrollDirectionConstants, ByVal shift As Integer, clientRectangle As PagerCtlLibACtl.RECTANGLE, ByVal horizontalScrollPosition As Long, ByVal verticalScrollPosition As Long, scrollDelta As Long)
  AddLogEntry "PagerA_BeforeScroll: scrollDirection=" & scrollDirection & ", shift=" & shift & ", clientRectangle=(" & clientRectangle.Left & "," & clientRectangle.Top & ")-(" & clientRectangle.Right & "," & clientRectangle.Bottom & "), horizontalScrollPosition=" & horizontalScrollPosition & ", verticalScrollPosition=" & verticalScrollPosition & ", scrollDelta=" & scrollDelta
End Sub

Private Sub PagerA_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  AddLogEntry "PagerA_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerA_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  AddLogEntry "PagerA_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerA_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  AddLogEntry "PagerA_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "PagerA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub PagerA_GetScrollableSize(ByVal Orientation As PagerCtlLibACtl.OrientationConstants, scrollableSize As Long)
  AddLogEntry "PagerA_GetScrollableSize: orientation=" & Orientation & ", scrollableSize=" & scrollableSize
  scrollableSize = Me.ScaleY(240 + 720 * 5 + Command1(0).Height + 240, vbTwips, vbPixels)
End Sub

Private Sub PagerA_GotFocus()
  AddLogEntry "PagerA_GotFocus"
  Set objActiveCtl = PagerA
End Sub

Private Sub PagerA_HotButtonChanging(ByVal scrollButton As PagerCtlLibACtl.ScrollButtonConstants, ByVal causedBy As PagerCtlLibACtl.HotButtonChangingCausedByConstants, ByVal additionalInfo As PagerCtlLibACtl.HotButtonChangingAdditionalInfoConstants, cancelChange As Boolean)
  AddLogEntry "PagerA_HotButtonChanging: scrollButton=" & scrollButton & ", causedBy=0x" & Hex(causedBy) & ", additionalInfo=0x" & Hex(additionalInfo) & ", cancelChange=" & cancelChange
End Sub

Private Sub PagerA_LostFocus()
  AddLogEntry "PagerA_LostFocus"
End Sub

Private Sub PagerA_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  AddLogEntry "PagerA_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerA_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  AddLogEntry "PagerA_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerA_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  AddLogEntry "PagerA_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerA_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  AddLogEntry "PagerA_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerA_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  AddLogEntry "PagerA_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerA_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  AddLogEntry "PagerA_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerA_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
'  AddLogEntry "PagerA_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerA_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  AddLogEntry "PagerA_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerA_OLEDragDrop(ByVal data As PagerCtlLibACtl.IOLEDataObject, effect As PagerCtlLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "PagerA_OLEDragDrop: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    PagerA.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub PagerA_OLEDragEnter(ByVal data As PagerCtlLibACtl.IOLEDataObject, effect As PagerCtlLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "PagerA_OLEDragEnter: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub PagerA_OLEDragLeave(ByVal data As PagerCtlLibACtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "PagerA_OLEDragLeave: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub PagerA_OLEDragMouseMove(ByVal data As PagerCtlLibACtl.IOLEDataObject, effect As PagerCtlLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "PagerA_OLEDragMouseMove: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub PagerA_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  AddLogEntry "PagerA_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerA_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  AddLogEntry "PagerA_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "PagerA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub PagerA_ResizedControlWindow()
  AddLogEntry "PagerA_ResizedControlWindow"
End Sub

Private Sub PagerA_Validate(Cancel As Boolean)
  AddLogEntry "PagerA_Validate"
End Sub

Private Sub PagerA_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  AddLogEntry "PagerA_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerA_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibACtl.HitTestConstants)
  AddLogEntry "PagerA_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerU_BeforeScroll(ByVal scrollDirection As PagerCtlLibUCtl.ScrollDirectionConstants, ByVal shift As Integer, clientRectangle As PagerCtlLibUCtl.RECTANGLE, ByVal horizontalScrollPosition As Long, ByVal verticalScrollPosition As Long, scrollDelta As Long)
  AddLogEntry "PagerU_BeforeScroll: scrollDirection=" & scrollDirection & ", shift=" & shift & ", clientRectangle=(" & clientRectangle.Left & "," & clientRectangle.Top & ")-(" & clientRectangle.Right & "," & clientRectangle.Bottom & "), horizontalScrollPosition=" & horizontalScrollPosition & ", verticalScrollPosition=" & verticalScrollPosition & ", scrollDelta=" & scrollDelta
End Sub

Private Sub PagerU_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  AddLogEntry "PagerU_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerU_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  AddLogEntry "PagerU_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerU_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  AddLogEntry "PagerU_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "PagerU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub PagerU_GetScrollableSize(ByVal Orientation As PagerCtlLibUCtl.OrientationConstants, scrollableSize As Long)
  AddLogEntry "PagerU_GetScrollableSize: orientation=" & Orientation & ", scrollableSize=" & scrollableSize
  scrollableSize = Me.ScaleY(240 + 720 * 5 + Command1(0).Height + 240, vbTwips, vbPixels)
End Sub

Private Sub PagerU_GotFocus()
  AddLogEntry "PagerU_GotFocus"
  Set objActiveCtl = PagerU
End Sub

Private Sub PagerU_HotButtonChanging(ByVal scrollButton As PagerCtlLibUCtl.ScrollButtonConstants, ByVal causedBy As PagerCtlLibUCtl.HotButtonChangingCausedByConstants, ByVal additionalInfo As PagerCtlLibUCtl.HotButtonChangingAdditionalInfoConstants, cancelChange As Boolean)
  AddLogEntry "PagerU_HotButtonChanging: scrollButton=" & scrollButton & ", causedBy=0x" & Hex(causedBy) & ", additionalInfo=0x" & Hex(additionalInfo) & ", cancelChange=" & cancelChange
End Sub

Private Sub PagerU_LostFocus()
  AddLogEntry "PagerU_LostFocus"
End Sub

Private Sub PagerU_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  AddLogEntry "PagerU_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerU_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  AddLogEntry "PagerU_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerU_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  AddLogEntry "PagerU_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerU_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  AddLogEntry "PagerU_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerU_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  AddLogEntry "PagerU_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerU_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  AddLogEntry "PagerU_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerU_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
'  AddLogEntry "PagerU_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerU_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  AddLogEntry "PagerU_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerU_OLEDragDrop(ByVal data As PagerCtlLibUCtl.IOLEDataObject, effect As PagerCtlLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "PagerU_OLEDragDrop: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    PagerU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub PagerU_OLEDragEnter(ByVal data As PagerCtlLibUCtl.IOLEDataObject, effect As PagerCtlLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "PagerU_OLEDragEnter: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub PagerU_OLEDragLeave(ByVal data As PagerCtlLibUCtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "PagerU_OLEDragLeave: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub PagerU_OLEDragMouseMove(ByVal data As PagerCtlLibUCtl.IOLEDataObject, effect As PagerCtlLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "PagerU_OLEDragMouseMove: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub PagerU_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  AddLogEntry "PagerU_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerU_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  AddLogEntry "PagerU_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "PagerU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub PagerU_ResizedControlWindow()
  AddLogEntry "PagerU_ResizedControlWindow"
End Sub

Private Sub PagerU_Validate(Cancel As Boolean)
  AddLogEntry "PagerU_Validate"
End Sub

Private Sub PagerU_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  AddLogEntry "PagerU_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub

Private Sub PagerU_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As PagerCtlLibUCtl.HitTestConstants)
  AddLogEntry "PagerU_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
End Sub


Private Sub AddLogEntry(ByVal txt As String)
  Dim pos As Long
  Static cLines As Long
  Static oldtxt As String

  If bLog Then
    cLines = cLines + 1
    If cLines > 50 Then
      ' delete the first line
      pos = InStr(oldtxt, vbNewLine)
      oldtxt = Mid(oldtxt, pos + Len(vbNewLine))
      cLines = cLines - 1
    End If
    oldtxt = oldtxt & (txt & vbNewLine)

    txtLog.Text = oldtxt
    txtLog.SelStart = Len(oldtxt)
  End If
End Sub

' returns the higher 16 bits of <value>
Private Function HiWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value) + LenB(ret), LenB(ret)

  HiWord = ret
End Function

' makes a 32 bits number out of two 16 bits parts
Private Function MakeDWord(ByVal Lo As Integer, ByVal hi As Integer) As Long
  Dim ret As Long

  CopyMemory VarPtr(ret), VarPtr(Lo), LenB(Lo)
  CopyMemory VarPtr(ret) + LenB(Lo), VarPtr(hi), LenB(hi)

  MakeDWord = ret
End Function

' subclasses this Form
Private Sub Subclass()
  Const NF_REQUERY = 4
  Const WM_NOTIFYFORMAT = &H55

  #If UseSubClassing Then
    If Not SubclassWindow(Me.hWnd, Me, EnumSubclassID.escidFrmMain) Then
      Debug.Print "Subclassing failed!"
    End If
    ' tell the control to negotiate the correct format with the form
    SendMessageAsLong PagerU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong PagerA.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
