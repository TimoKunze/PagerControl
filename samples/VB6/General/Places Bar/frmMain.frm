VERSION 5.00
Object = "{DF6596FB-AFA4-489C-94B0-FAC1615747AA}#1.2#0"; "PagerCtlU.ocx"
Object = "{5D0D0ABC-4898-4E46-AB48-291074A737A1}#1.1#0"; "TBarCtlsU.ocx"
Object = "{1F9B9092-BEE4-4CAF-9C7B-9384AF087C63}#1.4#0"; "ShBrowserCtlsU.ocx"
Object = "{9FC6639B-4237-4FB5-93B8-24049D39DF74}#1.5#0"; "ExLVwU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Pager 1.2 - Places Bar Sample"
   ClientHeight    =   6120
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   10215
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   681
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
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
      Left            =   3900
      TabIndex        =   3
      Top             =   5640
      Width           =   2415
   End
   Begin PagerCtlLibUCtl.Pager PagerU 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _cx             =   2778
      _cy             =   9551
      Appearance      =   1
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
      RightToLeftLayout=   0   'False
      ScrollAutomatically=   2
      ScrollButtonSize=   12
      SupportOLEDragImages=   -1  'True
      Begin TBarCtlsLibUCtl.ToolBar tbPlaces 
         Height          =   2655
         Left            =   240
         TabIndex        =   1
         Top             =   2280
         Width           =   780
         _cx             =   1376
         _cy             =   4683
         AllowCustomization=   0   'False
         AlwaysDisplayButtonText=   -1  'True
         AnchorHighlighting=   0   'False
         Appearance      =   0
         BackStyle       =   0
         BorderStyle     =   0
         ButtonHeight    =   64
         ButtonStyle     =   1
         ButtonTextPosition=   0
         ButtonWidth     =   96
         DisabledEvents  =   917739
         DisplayMenuDivider=   0   'False
         DisplayPartiallyClippedButtons=   -1  'True
         DontRedraw      =   0   'False
         DragClickTime   =   -1
         DragDropCustomizationModifierKey=   0
         DropDownGap     =   -1
         Enabled         =   -1  'True
         FirstButtonIndentation=   10
         FocusOnClick    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HighlightColor  =   -1
         HorizontalButtonPadding=   -1
         HorizontalButtonSpacing=   0
         HorizontalIconCaptionGap=   -1
         HorizontalTextAlignment=   1
         HoverTime       =   -1
         InsertMarkColor =   0
         MaximumButtonWidth=   0
         MaximumTextRows =   1
         MenuBarTheme    =   1
         MenuMode        =   0   'False
         MinimumButtonWidth=   0
         MousePointer    =   0
         MultiColumn     =   0   'False
         NormalDropDownButtonStyle=   0
         OLEDragImageStyle=   0
         Orientation     =   1
         ProcessContextMenuKeys=   -1  'True
         RaiseCustomDrawEventOnEraseBackground=   -1  'True
         RegisterForOLEDragDrop=   0
         RightToLeft     =   0
         ShadowColor     =   -1
         ShowToolTips    =   -1  'True
         SupportOLEDragImages=   -1  'True
         UseMnemonics    =   -1  'True
         UseSystemFont   =   -1  'True
         VerticalButtonPadding=   25
         VerticalButtonSpacing=   0
         VerticalTextAlignment=   0
         WrapButtons     =   0   'False
      End
   End
   Begin ExLVwLibUCtl.ExplorerListView ExLvw 
      Height          =   5400
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   8295
      _cx             =   14631
      _cy             =   9525
      AbsoluteBkImagePosition=   0   'False
      AllowHeaderDragDrop=   -1  'True
      AllowLabelEditing=   -1  'True
      AlwaysShowSelection=   -1  'True
      Appearance      =   1
      AutoArrangeItems=   2
      AutoSizeColumns =   0   'False
      BackColor       =   -2147483643
      BackgroundDrawMode=   0
      BkImagePositionX=   0
      BkImagePositionY=   0
      BkImageStyle    =   2
      BlendSelectionLasso=   -1  'True
      BorderSelect    =   0   'False
      BorderStyle     =   0
      CallBackMask    =   0
      CheckItemOnSelect=   0   'False
      ClickableColumnHeaders=   -1  'True
      ColumnHeaderVisibility=   1
      DisabledEvents  =   3145727
      DontRedraw      =   0   'False
      DragScrollTimeBase=   -1
      DrawImagesAsynchronously=   0   'False
      EditBackColor   =   -2147483643
      EditForeColor   =   -2147483640
      EditHoverTime   =   -1
      EditIMEMode     =   -1
      EmptyMarkupTextAlignment=   1
      Enabled         =   -1  'True
      FilterChangedTimeout=   -1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FullRowSelect   =   0
      GridLines       =   0   'False
      GroupFooterForeColor=   -2147483640
      GroupHeaderForeColor=   -2147483640
      GroupMarginBottom=   0
      GroupMarginLeft =   0
      GroupMarginRight=   0
      GroupMarginTop  =   12
      GroupSortOrder  =   0
      HeaderFullDragging=   -1  'True
      HeaderHotTracking=   0   'False
      HeaderHoverTime =   -1
      HeaderOLEDragImageStyle=   0
      HideLabels      =   0   'False
      HotForeColor    =   -1
      HotMousePointer =   0
      HotTracking     =   0   'False
      HotTrackingHoverTime=   -1
      HoverTime       =   -1
      IMEMode         =   -1
      IncludeHeaderInTabOrder=   0   'False
      InsertMarkColor =   0
      ItemActivationMode=   2
      ItemAlignment   =   0
      ItemBoundingBoxDefinition=   70
      ItemHeight      =   17
      JustifyIconColumns=   0   'False
      LabelWrap       =   -1  'True
      MinItemRowsVisibleInGroups=   0
      MousePointer    =   0
      MultiSelect     =   -1  'True
      OLEDragImageStyle=   1
      OutlineColor    =   -2147483633
      OwnerDrawn      =   0   'False
      ProcessContextMenuKeys=   -1  'True
      Regional        =   0   'False
      RegisterForOLEDragDrop=   -1  'True
      ResizableColumns=   -1  'True
      RightToLeft     =   0
      ScrollBars      =   1
      SelectedColumnBackColor=   -1
      ShowFilterBar   =   0   'False
      ShowGroups      =   0   'False
      ShowHeaderChevron=   0   'False
      ShowHeaderStateImages=   0   'False
      ShowStateImages =   0   'False
      ShowSubItemImages=   0   'False
      SimpleSelect    =   0   'False
      SingleRow       =   0   'False
      SnapToGrid      =   0   'False
      SortOrder       =   0
      SupportOLEDragImages=   -1  'True
      TextBackColor   =   -1
      TileViewItemLines=   2
      TileViewLabelMarginBottom=   0
      TileViewLabelMarginLeft=   0
      TileViewLabelMarginRight=   0
      TileViewLabelMarginTop=   0
      TileViewSubItemForeColor=   -1
      TileViewTileHeight=   -1
      TileViewTileWidth=   -1
      ToolTips        =   3
      UnderlinedItems =   0
      UseMinColumnWidths=   -1  'True
      UseSystemFont   =   -1  'True
      UseWorkAreas    =   0   'False
      View            =   0
      VirtualMode     =   0   'False
      EmptyMarkupText =   "frmMain.frx":0000
      FooterIntroText =   "frmMain.frx":006E
   End
   Begin ShBrowserCtlsLibUCtl.ShellListView SHLvw 
      Left            =   6600
      Top             =   5520
      Version         =   256
      AutoEditNewItems=   -1  'True
      AutoInsertColumns=   -1  'True
      ColorCompressedItems=   -1  'True
      ColorEncryptedItems=   -1  'True
      ColumnEnumerationTimeout=   3000
      DefaultManagedItemProperties=   511
      BeginProperty DefaultNamespaceEnumSettings {CC889E2B-5A0D-42F0-AA08-D5FD5863410C} 
         EnumerationFlags=   225
         ExcludedFileItemFileAttributes=   0
         ExcludedFileItemShellAttributes=   0
         ExcludedFolderItemFileAttributes=   0
         ExcludedFolderItemShellAttributes=   0
         IncludedFileItemFileAttributes=   0
         IncludedFileItemShellAttributes=   0
         IncludedFolderItemFileAttributes=   0
         IncludedFolderItemShellAttributes=   0
      EndProperty
      DisabledEvents  =   255
      DisplayElevationShieldOverlays=   -1  'True
      DisplayFileTypeOverlays=   -1
      DisplayThumbnailAdornments=   -1
      DisplayThumbnails=   0   'False
      HandleOLEDragDrop=   7
      HiddenItemsStyle=   2
      InfoTipFlags    =   -1073741824
      InitialSortOrder=   0
      ItemEnumerationTimeout=   3000
      ItemTypeSortOrder=   1
      LimitLabelEditInput=   -1  'True
      LoadOverlaysOnDemand=   -1  'True
      PersistColumnSettingsAcrossNamespaces=   1
      PreselectBasenameOnLabelEdit=   -1  'True
      ProcessShellNotifications=   -1  'True
      SelectSortColumn=   -1  'True
      SetSortArrows   =   -1  'True
      SortOnHeaderClick=   -1  'True
      ThumbnailSize   =   -1
      UseColumnResizability=   -1  'True
      UseFastButImpreciseItemHandling=   -1  'True
      UseGenericIcons =   1
      UseSystemImageList=   7
      UseThreadingForSlowColumns=   -1  'True
      UseThumbnailDiskCache=   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private Const MAX_PATH = 260


  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type

  Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName(1 To 2 * MAX_PATH) As Byte
    szTypeName(1 To 160) As Byte
  End Type


  Private bLog As Boolean
  Private themeableOS As Boolean


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
  Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function ILClone Lib "shell32.dll" Alias "#18" (ByVal pIDL As Long) As Long
  Private Declare Sub ILFree Lib "shell32.dll" Alias "#155" (ByVal pIDL As Long)
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
  Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long
  Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoW" (ByVal pszPath As Long, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
  Private Declare Function SHGetFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ppidl As Long) As Long


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


Private Sub cmdAbout_Click()
  PagerU.About
End Sub

Private Sub Form_Initialize()
  Dim hMod As Long

  InitCommonControls

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If
End Sub

Private Sub Form_Load()
  Const CSIDL_DESKTOP = &H0
  Const CSIDL_DRIVES = &H11
  Const CSIDL_FAVORITES = &H6
  Const CSIDL_PERSONAL = &H5
  Const CSIDL_RECENT = &H8
  Const CSIDL_CONTROLS = &H3
  Dim pIDL As Long
  Dim rc As RECT

  ' this is required to make the control work as expected
  Subclass
  
  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExLvw.hWnd, StrPtr("explorer"), 0
  End If
  PagerU.hContainedWindow = tbPlaces.hWnd
  
  GetClientRect PagerU.hWnd, rc
  tbPlaces.Width = ScaleX(rc.Right - rc.Left, vbPixels, vbTwips)
  tbPlaces.FirstButtonIndentation = (ScaleX(tbPlaces.Width, vbTwips, vbPixels) - tbPlaces.ButtonWidth) / 2
  
  SHLvw.Attach ExLvw.hWnd
  SHLvw.hWndShellUIParentWindow = Me.hWnd
  
  tbPlaces.LoadDefaultImages siltShellIcons48x48
  With tbPlaces.Buttons
    SHGetFolderLocation Me.hWnd, CSIDL_RECENT, 0, 0, pIDL
    .Add 1, , GetSysIconIndex(pIDL), , pIDLToDisplayName_Light(pIDL), , , btyCheckButton, True, , , pIDL, , , False, 96

    SHGetFolderLocation Me.hWnd, CSIDL_FAVORITES, 0, 0, pIDL
    .Add 2, , GetSysIconIndex(pIDL), , pIDLToDisplayName_Light(pIDL), , , btyCheckButton, True, , , pIDL, , , False, 96

    SHGetFolderLocation Me.hWnd, CSIDL_DESKTOP, 0, 0, pIDL
    .Add 3, , GetSysIconIndex(pIDL), , pIDLToDisplayName_Light(pIDL), , , btyCheckButton, True, , , pIDL, , , False, 96

    SHGetFolderLocation Me.hWnd, CSIDL_PERSONAL, 0, 0, pIDL
    .Add 4, , GetSysIconIndex(pIDL), , pIDLToDisplayName_Light(pIDL), , , btyCheckButton, True, , , pIDL, , , False, 96

    SHGetFolderLocation Me.hWnd, CSIDL_DRIVES, 0, 0, pIDL
    .Add 5, , GetSysIconIndex(pIDL), , pIDLToDisplayName_Light(pIDL), , , btyCheckButton, True, , , pIDL, , , False, 96

    SHGetFolderLocation Me.hWnd, CSIDL_CONTROLS, 0, 0, pIDL
    .Add 6, , GetSysIconIndex(pIDL), , pIDLToDisplayName_Light(pIDL), , , btyCheckButton, True, , , pIDL, , , False, 96
    
    .Item(3, btitID).SelectionState = ssChecked
    tbPlaces_ExecuteCommand 3, .Item(3, btitID), coButton, True
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim btn As ToolBarButton

  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
  For Each btn In tbPlaces.Buttons
    ILFree btn.ButtonData
  Next btn
  tbPlaces.Buttons.RemoveAll
End Sub

Private Sub PagerU_GetScrollableSize(ByVal Orientation As PagerCtlLibUCtl.OrientationConstants, scrollableSize As Long)
  scrollableSize = tbPlaces.IdealHeight
End Sub

Private Sub tbPlaces_CustomDraw(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, normalTextColor As stdole.OLE_COLOR, normalButtonBackColor As stdole.OLE_COLOR, normalBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, hotTextColor As stdole.OLE_COLOR, hotButtonBackColor As stdole.OLE_COLOR, markedTextBackColor As stdole.OLE_COLOR, markedButtonBackColor As stdole.OLE_COLOR, markedBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, ByVal drawStage As TBarCtlsLibUCtl.CustomDrawStageConstants, ByVal buttonState As TBarCtlsLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As TBarCtlsLibUCtl.RECTANGLE, textRectangle As TBarCtlsLibUCtl.RECTANGLE, HorizontalIconCaptionGap As stdole.OLE_XSIZE_PIXELS, furtherProcessing As TBarCtlsLibUCtl.CustomDrawReturnValuesConstants)
  Const COLOR_BTNHIGHLIGHT = 20
  Dim rc As RECT

  If drawStage = cdsPreErase Then
    ' If the application is not themed, we get drawing bugs, if we draw
    ' only the tool bar's or only the pager control's client area, so
    ' draw both of them.
    ' Note: If we'd draw using drawingRectangle, we'd get drawing bugs
    ' as soon as the control is scrolled.
    GetClientRect PagerU.hWnd, rc
    FillRect hDC, rc, GetSysColorBrush(COLOR_BTNHIGHLIGHT)
    GetClientRect tbPlaces.hWnd, rc
    'LSet rc = drawingRectangle
    FillRect hDC, rc, GetSysColorBrush(COLOR_BTNHIGHLIGHT)
    furtherProcessing = cdrvSkipDefault
  End If
End Sub

Private Sub tbPlaces_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  SHLvw.Namespaces.RemoveAll
  SHLvw.Namespaces.Add ILClone(toolButton.ButtonData)
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
    SendMessageAsLong ExLvw.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub

Private Function pIDLToDisplayName_Light(pIDLToDesktop As Long) As String
  Const SHGFI_DISPLAYNAME = &H200
  Const SHGFI_PIDL = &H8
  Dim Data As SHFILEINFO
  Dim ret As String

  ret = String$(MAX_PATH, 0)
  SHGetFileInfo pIDLToDesktop, 0, Data, LenB(Data), SHGFI_PIDL Or SHGFI_DISPLAYNAME
  lstrcpyn StrPtr(ret), VarPtr(Data.szDisplayName(1)), MAX_PATH
  pIDLToDisplayName_Light = Left$(ret, lstrlen(StrPtr(ret)))
End Function

Private Function GetSysIconIndex(ByVal pIDLToDesktop As Long) As Long
  Const SHGFI_PIDL = &H8
  Const SHGFI_SMALLICON = &H1
  Const SHGFI_SYSICONINDEX = &H4000
  Dim fileInfo As SHFILEINFO
  Dim Flags As Long

  Flags = SHGFI_PIDL Or SHGFI_SYSICONINDEX Or SHGFI_SMALLICON
  SHGetFileInfo pIDLToDesktop, 0, fileInfo, LenB(fileInfo), Flags
  GetSysIconIndex = fileInfo.iIcon
End Function
