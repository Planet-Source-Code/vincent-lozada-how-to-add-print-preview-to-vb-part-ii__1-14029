VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScrollBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ===========================================================================
' Name:     cScrollBars
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     24 December 1998
' Requires: SSUBTMR.DLL
'
' ---------------------------------------------------------------------------
' Copyright ® 1998 Steve McMahon (steve@vbaccelerator.com)
' Visit vbAccelerator - free, advanced source code for VB programmers.
'     http://vbaccelerator.com
' ---------------------------------------------------------------------------
'
' Description:
' A class which can add scroll bars to VB Forms, Picture Boxes and
' UserControls.
' Features:
'  * True API scroll bars, which don't flash or draw badly like
'    the VB ones
'  * Scroll bar values are long integers, i.e. >2 billion values
'  * Set Flat or Encarta scroll bar modes if your COMCTL32.DLL version
'    supports it (>4.72)
'
' FREE SOURCE CODE! - ENJOY.
' - Please report bugs to the author for incorporation into future releases
' - Don't sell this code.
' ===========================================================================

'private declare function InitializeFlatSB(HWND) as long
Private Declare Function InitialiseFlatSB _
        Lib "comctl32.dll" Alias "InitializeFlatSB" _
        ( _
        ByVal lHWnd As Long _
        ) As Long

' Memory functions:
Private Declare Sub CopyMemory _
        Lib "kernel32.dll" Alias "RtlMoveMemory" _
        ( _
        lpvDest As Any, _
        lpvSource As Any, _
        ByVal cbCopy As Long _
        )

' Window style bit functions:
Private Declare Function SetWindowLong _
        Lib "User32.dll" Alias "SetWindowLongA" _
        ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long _
        ) As Long

Private Declare Function GetWindowLong _
        Lib "User32.dll" Alias "GetWindowLongA" _
        ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long _
        ) As Long

' Window Long indexes:
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_HINSTANCE = (-6)
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_ID = (-12)
Private Const GWL_STYLE = (-16)
Private Const GWL_USERDATA = (-21)
Private Const GWL_WNDPROC = (-4)

' Creating new windows:
Private Declare Function CreateWindowEx _
        Lib "User32.dll" Alias "CreateWindowExA" _
        ( _
        ByVal dwExStyle As Long, _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String, _
        ByVal dwStyle As Long, _
        ByVal x As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hWndParent As Long, _
        ByVal hMenu As Long, _
        ByVal hInstance As Long, _
        lpParam As Any _
        ) As Long

' General window styles:
Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000    '  WS_BORDER Or WS_DLGFRAME
Private Const WS_CHILD = &H40000000
Private Const WS_CHILDWINDOW = (WS_CHILD)
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_DISABLED = &H8000000
Private Const WS_DLGFRAME = &H400000
Private Const WS_EX_ACCEPTFILES = &H10&
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_TRANSPARENT = &H20&
Private Const WS_GROUP = &H20000
Private Const WS_HSCROLL = &H100000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_ICONIC = WS_MINIMIZE
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_SYSMENU = &H80000
Private Const WS_TABSTOP = &H10000
Private Const WS_THICKFRAME = &H40000
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_VISIBLE = &H10000000
Private Const WS_VSCROLL = &H200000
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)

' Window appearance control:
Private Declare Function ShowWindow _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal nCmdShow As Long _
        ) As Long

' Show window styles
Private Const SW_SHOWNORMAL = 1
Private Const SW_ERASE = &H4
Private Const SW_HIDE = 0
Private Const SW_INVALIDATE = &H2
Private Const SW_MAX = 10
Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_NORMAL = 1
Private Const SW_OTHERUNZOOM = 4
Private Const SW_OTHERZOOM = 2
Private Const SW_PARENTCLOSING = 1
Private Const SW_RESTORE = 9
Private Const SW_PARENTOPENING = 3
Private Const SW_SHOW = 5
Private Const SW_SCROLLCHILDREN = &H1
Private Const SW_SHOWDEFAULT = 10
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_SHOWNOACTIVATE = 4
Private Declare Function SetWindowPos Lib "User32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_FRAMECHANGED = &H20    '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200    '  Don't do owner Z ordering
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const HWND_NOTOPMOST = -2

Private Declare Function DestroyWindow _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long _
        ) As Long

Private Declare Function EnableWindow _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal fEnable As Long _
        ) As Long

Private Declare Function SetFocus _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long _
        ) As Long

Private Declare Function MoveWindow _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal x As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal bRepaint As Long _
        ) As Long

Private Type RECT
    Left As Long
    Top As Long
    right As Long
    bottom As Long
End Type

Private Declare Function GetWindowRect _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        lpRECT As RECT _
        ) As Long

Private Declare Function InvalidateRect _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        lpRECT As RECT, _
        ByVal bErase As Long _
        ) As Long

' Window relationship functions:
Private Declare Function SetParent _
        Lib "User32.dll" _
        ( _
        ByVal hWndChild As Long, _
        ByVal hWndNewParent As Long _
        ) As Long

Private Declare Function GetParent _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long _
        ) As Long

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Declare Function ClientToScreen _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        lpPoint As POINTAPI _
        ) As Long

Private Declare Function ScreenToClient _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        lpPoint As POINTAPI _
        ) As Long

Private Declare Function ChildWindowFromPoint _
        Lib "User32.dll" _
        ( _
        ByVal hWndParent As Long, _
        ByVal x As Long, _
        ByVal Y As Long _
        ) As Long

Private Declare Function GetWindow _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal wCmd As Long _
        ) As Long

Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_MAX = 5
Private Const GW_OWNER = 4

Private Declare Function GetClientRECT _
        Lib "User32.dll" _
        Alias _
        "GetClientRect" (ByVal hWnd As _
        Long, lpRECT _
        As RECT) As Long

' Message functions:
Private Declare Function SendMessageByString _
        Lib "User32.dll" Alias "SendMessageA" _
        ( _
        ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As String _
        ) As Long

Private Declare Function SendMessageByLong _
        Lib "User32.dll" Alias "SendMessageA" _
        ( _
        ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long _
        ) As Long

Private Declare Function SendMessage _
        Lib "User32.dll" Alias "SendMessageA" _
        ( _
        ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any _
        ) As Long

Private Declare Function PostMessage _
        Lib "User32.dll" Alias "PostMessageA" _
        ( _
        ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long _
        ) As Long

' Key functions:
Private Declare Function GetAsyncKeyState _
        Lib "User32.dll" _
        ( _
        ByVal vKey As Long _
        ) As Integer

Private Declare Function GetCursorPos _
        Lib "User32.dll" _
        ( _
        lpPoint As POINTAPI _
        ) As Long

' Font
Private Const LF_FACESIZE = 32

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1

Private Declare Function CreateFontIndirect& _
        Lib "gdi32.dll" Alias "CreateFontIndirectA" _
        ( _
        lpLogFont As LOGFONT _
        )

Private Declare Function MulDiv _
        Lib "kernel32.dll" _
        ( _
        ByVal nNumber As Long, _
        ByVal nNumerator As Long, _
        ByVal nDenominator As Long _
        ) As Long

Private Declare Function SelectObject _
        Lib "gdi32.dll" _
        ( _
        ByVal hdc As Long, _
        ByVal hObject As Long _
        ) As Long

Private Declare Function DeleteObject _
        Lib "gdi32.dll" _
        ( _
        ByVal hObject As Long _
        ) As Long

Private Declare Function GetDeviceCaps _
        Lib "gdi32.dll" _
        ( _
        ByVal hdc As Long, _
        ByVal nIndex As Long _
        ) As Long

Private Const BITSPIXEL = 12
Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y

Private Declare Function GetSystemMetrics _
        Lib "User32.dll" _
        ( _
        ByVal nIndex As Long _
        ) As Long

Private Const SM_CXHSCROLL = 21
Private Const SM_CYHSCROLL = 3
Private Const SM_CYVSCROLL = 20
Private Const SM_CYVTHUMB = 9
Private Const SM_CXVSCROLL = 2
Private Const SM_CXHTHUMB = 10

' Scroll bar:
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Private Declare Function SetScrollInfo _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal n As Long, _
        lpcScrollInfo As SCROLLINFO, _
        ByVal BOOL As Boolean _
        ) As Long

Private Declare Function GetScrollInfo _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal n As Long, _
        LPSCROLLINFO As SCROLLINFO _
        ) As Long

Private Declare Function GetScrollPos _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal nBar As Long _
        ) As Long

Private Declare Function GetScrollRange _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal nBar As Long, _
        lpMinPos As Long, _
        lpMaxPos As Long _
        ) As Long

Private Declare Function SetScrollPos _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal nBar As Long, _
        ByVal nPos As Long, _
        ByVal bRedraw As Long _
        ) As Long

Private Declare Function SetScrollRange _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal nBar As Long, _
        ByVal nMinPos As Long, _
        ByVal nMaxPos As Long, _
        ByVal bRedraw As Long _
        ) As Long

Private Const SB_BOTH = 3
Private Const SB_BOTTOM = 7
Private Const SB_CTL = 2
Private Const SB_ENDSCROLL = 8
Private Const SB_HORZ = 0
Private Const SB_LEFT = 6
Private Const SB_LINEDOWN = 1
Private Const SB_LINELEFT = 0
Private Const SB_LINERIGHT = 1
Private Const SB_LINEUP = 0
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGELEFT = 2
Private Const SB_PAGERIGHT = 3
Private Const SB_PAGEUP = 2
Private Const SB_RIGHT = 7
Private Const SB_THUMBPOSITION = 4
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6
Private Const SB_VERT = 1

Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_DISABLENOSCROLL = &H8
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Private Const ESB_DISABLE_BOTH = &H3
Private Const ESB_ENABLE_BOTH = &H0

Private Declare Function EnableScrollBar _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal wSBflags As Long, _
        ByVal wArrows As Long _
        ) As Long

Private Declare Function ShowScrollBar _
        Lib "User32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal wBar As Long, _
        ByVal bShow As Long _
        ) As Long

Private Const WM_VSCROLL = &H115
Private Const WM_HSCROLL = &H114

' Flat scroll bars:
Private Const WSB_PROP_CYVSCROLL = &H1&
Private Const WSB_PROP_CXHSCROLL = &H2&
Private Const WSB_PROP_CYHSCROLL = &H4&
Private Const WSB_PROP_CXVSCROLL = &H8&
Private Const WSB_PROP_CXHTHUMB = &H10&
Private Const WSB_PROP_CYVTHUMB = &H20&
Private Const WSB_PROP_VBKGCOLOR = &H40&
Private Const WSB_PROP_HBKGCOLOR = &H80&
Private Const WSB_PROP_VSTYLE = &H100&
Private Const WSB_PROP_HSTYLE = &H200&
Private Const WSB_PROP_WINSTYLE = &H400&
Private Const WSB_PROP_PALETTE = &H800&
Private Const WSB_PROP_MASK = &HFFF&

Private Const FSB_FLAT_MODE = 2&
Private Const FSB_ENCARTA_MODE = 1&
Private Const FSB_REGULAR_MODE = 0&

Private Declare Function FlatSB_EnableScrollBar _
        Lib "comctl32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal int2 As Long, _
        ByVal UINT3 As Long _
        ) As Long

Private Declare Function FlatSB_ShowScrollBar _
        Lib "comctl32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal code As Long, _
        ByVal fRedraw As Boolean _
        ) As Long

Private Declare Function FlatSB_GetScrollRange _
        Lib "comctl32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal code As Long, _
        ByVal LPINT1 As Long, _
        ByVal LPINT2 As Long _
        ) As Long

Private Declare Function FlatSB_GetScrollInfo _
        Lib "comctl32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal code As Long, _
        LPSCROLLINFO As SCROLLINFO _
        ) As Long

Private Declare Function FlatSB_GetScrollPos _
        Lib "comctl32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal code As Long _
        ) As Long

Private Declare Function FlatSB_GetScrollProp _
        Lib "comctl32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal propIndex As Long, _
        ByVal LPINT As Long _
        ) As Long

Private Declare Function FlatSB_SetScrollPos _
        Lib "comctl32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal code As Long, _
        ByVal pos As Long, _
        ByVal fRedraw As Boolean _
        ) As Long

Private Declare Function FlatSB_SetScrollInfo _
        Lib "comctl32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal code As Long, _
        LPSCROLLINFO As SCROLLINFO, _
        ByVal fRedraw As Boolean _
        ) As Long

Private Declare Function FlatSB_SetScrollRange _
        Lib "comctl32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal code As Long, _
        ByVal Min As Long, _
        ByVal Max As Long, _
        ByVal fRedraw As Boolean _
        ) As Long

Private Declare Function FlatSB_SetScrollProp _
        Lib "comctl32.dll" _
        ( _
        ByVal hWnd As Long, _
        ByVal Index As Long, _
        ByVal newValue As Long, _
        ByVal fRedraw As Boolean _
        ) As Long

Private Declare Function InitializeFlatSB _
        Lib "comctl32.dll" _
        ( _
        ByVal hWnd As Long _
        ) As Long

Private Declare Function UninitializeFlatSB _
        Lib "comctl32.dll" _
        ( _
        ByVal hWnd As Long _
        ) As Long

' Message response:
Implements ISubclass
Private m_emr As EMsgResponse

' Initialisation state:
Private m_bInitialised As Boolean

' Orientation
Public Enum FSOrientationConstants
    fsoHorizontal
    fsoVertical
    fsoBoth
End Enum

Private m_eOrientation As FSOrientationConstants

' Style
Public Enum FSStyleConstants
    fsRegular = FSB_REGULAR_MODE
    fsEncarta = FSB_ENCARTA_MODE
    fsFlat = FSB_FLAT_MODE
End Enum

Private m_eStyle As FSStyleConstants

' Bars:
Public Enum FSScrollBarConstants
    fsHorizontal = SB_HORZ
    fsVertical = SB_VERT
End Enum

' Can we have flat scroll bars?
Private m_bNoFlatScrollBars As Boolean

' hWnd we're adding scroll bars too:
Private m_hWnd As Long

' Small change amount
Private m_lSmallChangeHorz As Long
Private m_lSmallChangeVert As Long
' Enabled:
Private m_bEnabledHorz As Boolean
Private m_bEnabledVert As Boolean
' Visible
Private m_bVisibleHorz As Boolean
Private m_bVisibleVert As Boolean
Private m_bScroll As Boolean

Public Event Scroll(eBar As FSScrollBarConstants)
Public Event Change(eBar As FSScrollBarConstants)

Public Property Get Visible( _
    ByVal eBar As FSScrollBarConstants _
    ) As Boolean

    If (eBar = fsHorizontal) Then
        Visible = m_bVisibleHorz
    Else
        Visible = m_bVisibleVert
    End If
End Property
Public Property Let Visible( _
    ByVal eBar As FSScrollBarConstants, _
    ByVal bState As Boolean)

    If (eBar = fsHorizontal) Then
        m_bVisibleHorz = bState
        If Not bState Then
            LargeChange(fsHorizontal) = 1
            SmallChange(fsHorizontal) = 1
            Max(fsHorizontal) = 1
        End If
    Else
        m_bVisibleVert = bState
        If Not bState Then
            LargeChange(fsVertical) = 1
            SmallChange(fsVertical) = 1
            Max(fsVertical) = 1
        End If
    End If
    If (m_bNoFlatScrollBars) Then
        Call ShowScrollBar(m_hWnd, eBar, Abs(bState))
    Else
        Call FlatSB_ShowScrollBar(m_hWnd, eBar, Abs(bState))
    End If
End Property

Public Property Get Orientation() As FSOrientationConstants
    Orientation = m_eOrientation
End Property

Public Property Let Orientation( _
    ByVal eOrientation As FSOrientationConstants)

    m_eOrientation = eOrientation
    pSetOrientation
End Property

Private Sub pSetOrientation()
    Call ShowScrollBar(m_hWnd, _
            SB_HORZ, _
            Abs((m_eOrientation = fsoBoth) Or (m_eOrientation = fsoHorizontal)))
    Call ShowScrollBar(m_hWnd, _
            SB_VERT, _
            Abs((m_eOrientation = fsoBoth) Or (m_eOrientation = fsoVertical)))
End Sub

Private Sub pGetSI( _
    ByVal eBar As FSScrollBarConstants, _
    ByRef tSI As SCROLLINFO, _
    ByVal fMask As Long)
    Dim lO As Long

    lO = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)

    If (m_bNoFlatScrollBars) Then
        GetScrollInfo m_hWnd, lO, tSI
    Else
        FlatSB_GetScrollInfo m_hWnd, lO, tSI
    End If
End Sub

Private Sub pLetSI( _
    ByVal eBar As FSScrollBarConstants, _
    ByRef tSI As SCROLLINFO, _
    ByVal fMask As Long)
    Dim lO As Long

    lO = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)

    If (m_bNoFlatScrollBars) Then
        SetScrollInfo m_hWnd, lO, tSI, True
    Else
        FlatSB_SetScrollInfo m_hWnd, lO, tSI, True
    End If
End Sub

Property Get Style() As FSStyleConstants
    Style = m_eStyle
End Property

Property Let Style( _
    ByVal eStyle As FSStyleConstants)
    Dim lR As Long

    If (eStyle <> fsRegular) Then
        If (m_bNoFlatScrollBars) Then
            ' can't do it..
            Exit Property
        End If
    End If

    If (m_eOrientation = fsoHorizontal) Or (m_eOrientation = fsoBoth) Then
        lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_HSTYLE, eStyle, True)
    End If

    If (m_eOrientation = fsoVertical) Or (m_eOrientation = fsoBoth) Then
        lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_VSTYLE, eStyle, True)
    End If

    m_eStyle = eStyle
End Property

Public Property Get SmallChange( _
    ByVal eBar As FSScrollBarConstants _
    ) As Long

    If (eBar = fsHorizontal) Then
        SmallChange = m_lSmallChangeHorz
    Else
        SmallChange = m_lSmallChangeVert
    End If
End Property

Property Let SmallChange( _
    ByVal eBar As FSScrollBarConstants, _
    ByVal lSmallChange As Long)

    If (eBar = fsHorizontal) Then
        m_lSmallChangeHorz = lSmallChange
    Else
        m_lSmallChangeVert = lSmallChange
    End If
End Property

Property Get Enabled( _
    ByVal eBar As FSScrollBarConstants _
    ) As Boolean

    If (eBar = fsHorizontal) Then
        Enabled = m_bEnabledHorz
    Else
        Enabled = m_bEnabledVert
    End If
End Property

Property Let Enabled( _
    ByVal eBar As FSScrollBarConstants, _
    ByVal bEnabled As Boolean)
    Dim lO As Long
    Dim lF As Long

    lO = eBar
    If (bEnabled) Then
        lF = ESB_ENABLE_BOTH
    Else
        lF = ESB_DISABLE_BOTH
    End If

    If (m_bNoFlatScrollBars) Then
        Call EnableScrollBar(m_hWnd, lO, lF)
    Else
        Call FlatSB_EnableScrollBar(m_hWnd, lO, lF)
    End If
End Property

Property Get Min( _
    ByVal eBar As FSScrollBarConstants _
    ) As Long
    Dim tSI As SCROLLINFO

    pGetSI eBar, tSI, SIF_RANGE
    Min = tSI.nMin
End Property

Property Get Max( _
    ByVal eBar As FSScrollBarConstants _
    ) As Long
    Dim tSI As SCROLLINFO

    Call pGetSI(eBar, tSI, SIF_RANGE Or SIF_PAGE)
    Max = tSI.nMax - tSI.nPage
End Property

Property Get Value( _
    ByVal eBar As FSScrollBarConstants _
    ) As Long
    Dim tSI As SCROLLINFO

    Call pGetSI(eBar, tSI, SIF_POS)
    Value = tSI.nPos
End Property

Property Get LargeChange( _
    ByVal eBar As FSScrollBarConstants _
    ) As Long
    Dim tSI As SCROLLINFO

    Call pGetSI(eBar, tSI, SIF_PAGE)
    LargeChange = tSI.nPage
End Property

Property Let Min( _
    ByVal eBar As FSScrollBarConstants, _
    ByVal iMin As Long)
    Dim tSI As SCROLLINFO

    tSI.nMin = iMin
    tSI.nMax = Max(eBar) + LargeChange(eBar)
    Call pLetSI(eBar, tSI, SIF_RANGE)
End Property

Property Let Max( _
    ByVal eBar As FSScrollBarConstants, _
    ByVal iMax As Long)
    Dim tSI As SCROLLINFO

    tSI.nMax = iMax + LargeChange(eBar)
    tSI.nMin = Min(eBar)
    Call pLetSI(eBar, tSI, SIF_RANGE)
    RaiseEvent Change(eBar)
End Property

Property Let Value( _
    ByVal eBar As FSScrollBarConstants, _
    ByVal iValue As Long)
    Dim tSI As SCROLLINFO

    If (iValue <> Value(eBar)) And m_bScroll = False Then
        tSI.nPos = iValue
        Call pLetSI(eBar, tSI, SIF_POS)
        RaiseEvent Change(eBar)
    ElseIf (iValue <> Value(eBar)) And m_bScroll = True Then
        tSI.nPos = iValue
        Call pLetSI(eBar, tSI, SIF_POS)
    End If
End Property

Property Let LargeChange( _
    ByVal eBar As FSScrollBarConstants, _
    ByVal iLargeChange As Long)
    Dim tSI As SCROLLINFO
    Dim lCurMax As Long
    Dim lCurLargeChange As Long

    Call pGetSI(eBar, tSI, SIF_ALL)
    tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
    tSI.nPage = iLargeChange
    Call pLetSI(eBar, tSI, SIF_PAGE Or SIF_RANGE)
End Property

Property Get CanBeFlat() As Boolean
    CanBeFlat = Not (m_bNoFlatScrollBars)
End Property

Private Sub pCreateScrollBar()
    Dim lR As Long
    Dim lStyle As Long
    Dim hParent As Long

    ' Redraw the window so it will take (pinched from MFC):
    hParent = GetParent(m_hWnd)

    Call SetParent(m_hWnd, hParent)
    ' Redraw for added insurance ('Slimy Windows Hacks')
    Call SetWindowPos(m_hWnd, _
            HWND_NOTOPMOST, _
            0, _
            0, _
            0, _
            0, _
            SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOMOVE Or SWP_DRAWFRAME)

    On Error Resume Next
    lR = InitialiseFlatSB(m_hWnd)

    If (Err.Number <> 0) Then
        'Can't find DLL entry point InitializeFlatSB in COMCTL32.DLL
        ' Means we have version prior to 4.71
        ' We get standard scroll bars.
        m_bNoFlatScrollBars = True
    Else
        Style = m_eStyle
    End If

End Sub

Public Sub Create( _
    ByVal hWndA As Long)

    Call pClearUp
    m_hWnd = hWndA
    Call pCreateScrollBar
    Call pAttachMessages
End Sub

Private Sub pClearUp()
    If m_hWnd <> 0 Then
        On Error Resume Next
        ' Stop flat scroll bar if we have it:
        If Not (m_bNoFlatScrollBars) Then
            Call UninitializeFlatSB(m_hWnd)
        End If

        On Error GoTo 0
        ' Remove subclass:
        Call DetachMessage(Me, m_hWnd, WM_HSCROLL)
        Call DetachMessage(Me, m_hWnd, WM_VSCROLL)
    End If

    m_hWnd = 0
    m_bInitialised = False
End Sub

Private Sub pAttachMessages()
    If (m_hWnd <> 0) Then
        Call AttachMessage(Me, m_hWnd, WM_HSCROLL)
        Call AttachMessage(Me, m_hWnd, WM_VSCROLL)
        m_bInitialised = True
    End If
End Sub

Private Sub Class_Initialize()
    m_lSmallChangeHorz = 1
    m_lSmallChangeVert = 1
    m_eStyle = fsRegular
    m_eOrientation = fsoBoth
End Sub

Private Sub Class_Terminate()
    Call pClearUp
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer.EMsgResponse)
    'Nothing
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer.EMsgResponse
    ISubclass_MsgResponse = emrPostProcess
End Property

Private Function ISubclass_WindowProc( _
    ByVal hWnd As Long, _
    ByVal iMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
    ) As Long
    Dim lScrollCode As Long
    Dim tSI As SCROLLINFO
    Dim lV As Long, lSC As Long
    Dim eBar As FSScrollBarConstants

    Select Case iMsg
        Case WM_VSCROLL, WM_HSCROLL
            If (iMsg = WM_HSCROLL) Then
                eBar = fsHorizontal
            Else
                eBar = fsVertical
            End If

            lScrollCode = (wParam And &HFFFF&)

            Select Case lScrollCode
                Case SB_THUMBTRACK
                    'Is vertical/horizontal?
                    Call pGetSI(eBar, tSI, SIF_TRACKPOS)
                    m_bScroll = True
                    Value(eBar) = tSI.nTrackPos
                    m_bScroll = False
                    RaiseEvent Scroll(eBar)
                Case SB_LEFT, SB_BOTTOM
                    Value(eBar) = Min(eBar)
                    RaiseEvent Change(eBar)
                Case SB_RIGHT, SB_TOP
                    Value(eBar) = Max(eBar)
                    RaiseEvent Change(eBar)
                Case SB_LINELEFT, SB_LINEUP
                    lV = Value(eBar)
                    If (eBar = fsHorizontal) Then
                        lSC = m_lSmallChangeHorz
                    Else
                        lSC = m_lSmallChangeVert
                    End If

                    If (lV - lSC < Min(eBar)) Then
                        Value(eBar) = Min(eBar)
                    Else
                        Value(eBar) = lV - lSC
                    End If
                    RaiseEvent Change(eBar)

                Case SB_LINERIGHT, SB_LINEDOWN
                    lV = Value(eBar)
                    If (eBar = fsHorizontal) Then
                        lSC = m_lSmallChangeHorz
                    Else
                        lSC = m_lSmallChangeVert
                    End If

                    If (lV + lSC > Max(eBar)) Then
                        Value(eBar) = Max(eBar)
                    Else
                        Value(eBar) = lV + lSC
                    End If
                    RaiseEvent Change(eBar)
                Case SB_PAGELEFT, SB_PAGEUP
                    Value(eBar) = Value(eBar) - LargeChange(eBar)
                    RaiseEvent Change(eBar)
                Case SB_PAGERIGHT, SB_PAGEDOWN
                    Value(eBar) = Value(eBar) + LargeChange(eBar)
                    RaiseEvent Change(eBar)
                Case SB_ENDSCROLL
                    RaiseEvent Change(eBar)
            End Select
    End Select
End Function

'Private Function pRaiseEvent( _
 '    ByVal eBar As FSScrollBarConstants, _
 '    ByVal bScroll As Boolean)
'    Static s_lLastValue(0 To 1) As Long
'
'    If (Value(eBar) <> s_lLastValue(eBar)) Then
'        If (bScroll) Then
'            RaiseEvent Scroll(eBar)
'        Else
'            RaiseEvent Change(eBar)
'        End If
'        s_lLastValue(eBar) = Value(eBar)
'    End If
'End Function
