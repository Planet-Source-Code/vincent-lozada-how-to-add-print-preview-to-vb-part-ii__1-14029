VERSION 5.00
Begin VB.UserControl Preview 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   EditAtDesignTime=   -1  'True
   Enabled         =   0   'False
   PropertyPages   =   "PageView.ctx":0000
   ScaleHeight     =   4995
   ScaleWidth      =   6000
   ToolboxBitmap   =   "PageView.ctx":0035
   Begin VB.PictureBox picWorkspace 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4695
      ScaleWidth      =   5775
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.PictureBox picPage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   1440
         ScaleHeight     =   3735
         ScaleWidth      =   2775
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   2775
         Begin VB.Shape shpPaperBorder 
            BorderStyle     =   6  'Inside Solid
            Height          =   3255
            Left            =   240
            Top             =   240
            Width           =   2265
         End
      End
      Begin VB.PictureBox picPageShadow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   3735
         Left            =   1560
         ScaleHeight     =   3705
         ScaleWidth      =   2745
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.Image imgNoPrinters 
         Appearance      =   0  'Flat
         Height          =   720
         Left            =   360
         Picture         =   "PageView.ctx":0347
         Top             =   1320
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblNoPrinters 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No printers are installed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Visible         =   0   'False
         Width           =   1305
      End
   End
End
Attribute VB_Name = "Preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'*******************************************************************************************
' Class:    PageView
' Filename: PageView.cls
' Author:   Vincent Lozada
' Date:     04/17/2000
'*******************************************************************************************

Option Explicit

'===============================================
' PUBLIC DECLARATIONS
'===============================================
Public Pages    As Pages
Public Header   As PageHeaderFooter
Public Footer   As PageHeaderFooter

'===============================================
' PUBLIC CONSTANT DECLARATIONS
'===============================================
Public Enum AbortWindowPositionConstants
    awpAppWindow = 0    'Causes the automatic Abort dialog to appear centered over the vsPrinter control.
    awpScreenCenter = 1    'Causes the automatic Abort dialog to appear centered on the screen.
End Enum

Public Enum ActionConstants
    paNone = 0    'No effect.
    paPrintFile = 1    'Prints the file specified by the FileName property.
    paChoosePrintFile = 2    'Shows a printer setup dialog, then prints the file specified by the FileName property.
    paStartDoc = 3    'Equivalent to the StartDoc method.
    paNewPage = 4    'Equivalent to the NewPage method.
    paNewCol = 5    'Equivalent to the NewColumn method.
    paEndDoc = 6    'Equivalent to the EndDoc method.
    paAbortDoc = 7    'Equivalent to the KillDoc method.
    paPrintPage = 8    'Prints the current preview page (set by the PreviewPage property).
    paChoosePrintPage = 9    'Shows a printer setup dialog, then prints the current preview page.
    paCopyPage = 10    'Copies the current preview page to the clipboard.
    paPrintAll = 11    'Equivalent to the PrintDoc method.
    paChoosePrintAll = 12    'Shows a printer setup dialog, then prints the document being previewed.
    paChoosePrinter = 13    'Equivalent to the PrintDialog method.
    paPageSetup = 14    'Equivalent to the PrintDialog method.
End Enum

Public Enum AppearanceConstants
    apFlat = 0
    ap3D = 1
End Enum

Public Enum BorderStyleConstants
    bsNone = 0
    bsSingle = 1
End Enum

Public Enum ColorModeConstants
    cmMonochrome = 1    'Print output in monochrome (usually shades of black and white).
    cmColor = 2    'Print output in color.
End Enum

Public Enum CollateConstants
    colFalse = 0    'Do not collate when printing multiple copies.
    colTrue = 1    'Collate when printing multiple copies.
End Enum

Public Enum DrawStyleConstants
    dsSolid = 0    '(Default) Solid
    dsDash = 1    'Dash
    dsDot = 2    'Dot
    dsDashDot = 3    'Dash-Dot
    dsDashDotDot = 4    'Dash-Dot-Dot
    dsInvisible = 5    'Transparent
    dsInsideSolid = 6    'Inside Solid
End Enum

Public Enum DuplexConstants
    dpxSimplex = 1    'Single-sided printing with the current orientation setting.
    dpxHorizontal = 2    'Double-sided printing using a horizontal page turn.
    dpxVertical = 3    'Double-sided printing using a vertical page turn.
End Enum

Public Enum FillStyleConstants
    fsSolid = 0    'Solid
    fsTransparent = 1    '(Default) Transparent
    fsHorizontalLine = 2    'Horizontal Line
    fsVerticalLine = 3    'Vertical Line
    fsUpwardDiagonal = 4    'Upward Diagonal
    fsDownwardDiagonal = 5    'Downward Diagonal
    fsCross = 6    'Cross
    fsDiagonalCross = 7    'Diagonal Cross
End Enum

Public Enum MousePointerConstants
    mpDefault = 0    'Default
    mpArrow = 1    'Arrow
    mpCrosshair = 2    'Cross
    mpIbeam = 3    'I beam
    mpIconPointer = 4    'Icon
    mpSizePointer = 5    'Size
    mpSizeNESW = 6    'Size NE, SW
    mpSizeNS = 7    'Size N, S
    mpSizeNWSE = 8    'Size NW, SE
    mpSizeWE = 9    'Size W, E
    mpUpArrow = 10    'Up arrow
    mpHourglass = 11    'Hourglass
    mpNoDrop = 12    'No drop
    mpArrowHourglass = 13    'Arrow and hourglass; (available only in 32-bit Visual Basic 5.0)
    mpArrowQuestion = 14    'Arrow and question mark; (available only in 32-bit Visual Basic 5.0)
    mpSizeAll = 15    'Size all; (available only in 32-bit Visual Basic 5.0)
    mpCustom = 99    'Custom icon specified by the MouseIcon property
End Enum

Public Enum MouseZoomConstants
    mzNone = 0    'No automatic zooming with the mouse.
    mzSimple = 1    'Clicking the mouse with the left- and right-buttons zooms in and out
    'by incrementing or decrementing the Zoom property by a step specified by the
    'ZoomStep property. The ZoomMode property is alway set to zmPercentage.
    mzExtended = 2    'Clicking the mouse with the left- and right-buttons zooms in and out
    'by incrementing or decrementing the Zoom property by a step specified by the
    'ZoomStep property. If the user zooms out until the entire preview page is
    'visible, the control switches the ZoomMode property to show two or more pages
    'at once.
End Enum

Public Enum OrientationConstants
    orPortrait = 1    'Documents are printed with the top at the narrow side of the paper.
    orLandscape = 2    'Documents are printed with the top at the wide side of the paper.
End Enum

Public Enum PageBorderConstants
    pbNone = 0
    pbBottom = 1
    pbTop = 2
    pbTopBottom = 3
    pbBox = 4
End Enum

Public Enum PaperBinConstants
    pbUpper = 1    'Use paper from the upper bin.
    pbLower = 2    'Use paper from the lower bin.
    pbMiddle = 3    'Use paper from the middle bin.
    pbManual = 4    'Wait for manual insertion of each sheet of paper.
    pbEnvelope = 5    'Use envelopes from the envelope feeder.
    pbEnvManual = 6    'Use envelopes from the envelope feeder, but wait for manual insertion.
    pbAuto = 7    '(Default) Use paper from the current default bin.
    pbTractor = 8    'Use paper fed from the tractor feeder.
    pbSmallFmt = 9    'Use paper from the small paper feeder.
    pbLargeFmt = 10    'Use paper from the large paper bin.
    pbLargeCapacity = 11    'Use paper from the large capacity feeder.
    pbCassette = 14    'Use paper from the attached cassette cartridge.
End Enum

Public Enum PaperSizeConstants
    psLetter = 1    'Letter, 8 1/2 x 11 in.
    psLetterSmall = 2    'Letter Small, 8 1/2 x 11 in.
    psTabloid = 3    'Tabloid, 11 x 17 in.
    psLedger = 4    'Ledger, 17 x 11 in.
    psLegal = 5    'Legal, 8 1/2 x 14 in.
    psStatement = 6    'Statement, 5 1/2 x 8 1/2 in.
    psExecutive = 7    'Executive, 7 1/2 x 10 1/2 in.
    psA3 = 8    'A3, 297 x 420 mm
    psA4 = 9    'A4, 210 x 297 mm
    psA4Small = 10    'A4 Small, 210 x 297 mm
    psA5 = 11    'A5, 148 x 210 mm
    psB4 = 12    'B4, 250 x 354 mm
    psB5 = 13    'B5, 182 x 257 mm
    psFolio = 14    'Folio, 8 1/2 x 13 in.
    psQuarto = 15    'Quarto, 215 x 275 mm
    ps10x14 = 16    '10 x 14 in.
    ps11x17 = 17    '11 x 17 in.
    psNote = 18    'Note, 8 1/2 x 11 in.
    psEnv9 = 19    'Envelope #9, 3 7/8 x 8 7/8 in.
    psEnv10 = 20    'Envelope #10, 4 1/8 x 9 1/2 in.
    psEnv11 = 21    'Envelope #11, 4 1/2 x 10 3/8 in.
    psEnv12 = 22    'Envelope #12, 4 1/2 x 11 in.
    psEnv14 = 23    'Envelope #14, 5 x 11 1/2 in.
    psCSheet = 24    'C size sheet
    psDSheet = 25    'D size sheet
    psESheet = 26    'E size sheet
    psEnvDL = 27    'Envelope DL, 110 x 220 mm
    psEnvC3 = 29    'Envelope C3, 324 x 458 mm
    psEnvC4 = 30    'Envelope C4, 229 x 324 mm
    psEnvC5 = 28    'Envelope C5, 162 x 229 mm
    psEnvC6 = 31    'Envelope C6, 114 x 162 mm
    psEnvC65 = 32    'Envelope C65, 114 x 229 mm
    psEnvB4 = 33    'Envelope B4, 250 x 353 mm
    psEnvB5 = 34    'Envelope B5, 176 x 250 mm
    psEnvB6 = 35    'Envelope B6, 176 x 125 mm
    psEnvItaly = 36    'Envelope, 110 x 230 mm
    psEnvMonarch = 37    'Envelope Monarch, 3 7/8 x 7 1/2 in.
    psEnvPersonal = 38    'Envelope, 3 5/8 x 6 1/2 in.
    psFanfoldUS = 39    'U.S. Standard Fanfold, 14 7/8 x 11 in.
    psFanfoldStdGerman = 40    'German Standard Fanfold, 8 1/2 x 12 in.
    psFanfoldLglGerman = 41    'German Legal Fanfold, 8 1/2 x 13 in.
    psUser = 256    'User-defined
End Enum

Public Enum PrintDialogConstants
    pdPrinterSetup = 0    'Display the standard printer selection dialog, which allows you to select a printer and modify its properties.
    pdPageSetup = 1    'Display the standard page setup dialog, which allows you to set margins, paper size, and page orientation.
End Enum

Public Enum PrintQualityConstants
    pqDraft = -1    'Draft resolution
    pqLow = -2    'Low resolution
    pqMedium = -3    'Medium resolution
    pqHigh = -4    'High resolution
End Enum

Public Enum PrinterErrorConstants
    perrNone = 0    'No error.
    perrCantAccessPrinter = 3    'Printer is not available.
    perrCantStartJob = 4    'Printer is not responding.
    perrUserAborted = 5    'User clicked the CANCEL button in the automatic Abort dialog.
    perrAlreadyPrinting = 6    'Can't start a new document while another document is being created.
    perrDeviceIncapable = 7    'Printer driver does not support the requested setting.
    perrControlIncapable = 8    'Trying to use the RenderControl property with a control that does not support OPP.
    perrCantInBrowser = 9    'Can't access disk when running in browser safety mode.
End Enum

Public Enum ShapeTypeConstants
    stRectangle = 0
    stSquare = 1
    stOval = 2
    stCircle = 3
    stRoundedRectangle = 4
    stRoundedSquare = 5
    stLine = 6
End Enum

Public Enum ShapeBorderStyleConstants
    sbsTransparent = 0
    sbsSolid = 1
    sbsDash = 2
    sbsDot = 3
    sbsDashDot = 4
    sbsDashDotDot = 5
    sbsInsideSolid = 6
End Enum

Public Enum ShowGuidesConstants
    sgHide = 0    'Never show guides.
    sgShow = 1    'Always show guides.
    sgDesignTime = 2    'Show guides at design time (default setting).
End Enum

Public Enum TextAlignmentConstants
    taLeftTop = 0    'Text is aligned to the left and to the top.
    taCenterTop = 1    'Text is aligned to the center and to the top.
    taRightTop = 2    'Text is aligned to the right and to the top.
    taLeftBottom = 3    'Text is aligned to the left and to the bottom.
    taCenterBottom = 4    'Text is aligned to the center and to the bottom.
    taRightBottom = 5    'Text is aligned to the right and to the bottom.
    taLeftMiddle = 6    'Text is aligned to the left and to the middle.
    taCenterMiddle = 7    'Text is aligned to the center and to the middle.
    taRightMiddle = 8    'Text is aligned to the right and to the middle.
    taJustTop = 9    'Text is fully justified and aligned to the top.
    taJustBottom = 10    'Text is fully justified and aligned to the bottom.
    taJustMiddle = 11    'Text is fully justified and aligned to the middle.
End Enum

Public Enum TrueTypeFontsConstants
    ttfBitmap = 1    'Prints TrueType fonts as raster graphics.
    ttfDownload = 2    'Downloads TrueType fonts as soft fonts.
    ttfSubDevice = 3    'Substitute device fonts for TrueType fonts.
    ttfOutline = 4    'Prints TrueType fonts as vector graphics.
End Enum

Public Enum ZoomModeConstants
    zmPercentage = 0    'Use zoom factor set by the user with Zoom property.
    zmThumbnail = 1    'Show as many 1-inch wide preview pages as will fit on the control.
    zmTwoPages = 2    'Show two whole pages, side by side.
    zmWholePage = 3    'Show a whole page.
    zmPageWidth = 4    'Show the page so that it fits horizontally within the control.
    zmStretch = 5    'Stretch the page to fill the control without preserving the aspect ratio.
End Enum

Public Type ABORTWINDOWTYPE
    str_WindowCaption       As String
    str_ButtonCaption       As String
    str_DeviceNameCaption   As String
    str_PageCaption         As String
    enum_WindowPosition     As AbortWindowPositionConstants
End Type

'===============================================
' PRIVATE DECLARATIONS
'===============================================
Private WithEvents ScrollBars As cScrollBars
Attribute ScrollBars.VB_VarHelpID = -1

Private m_enum_CurrentZoomScale     As ZoomScaleConstants
Private Enum ZoomScaleConstants
    zsWholePage = 0
    zs100 = 1
End Enum

'Constants
Private Const m_cint_WkspPgMargin   As Integer = 100
Private Const m_cint_TwipsPerInch   As Integer = 1440
Private Const m_clng_Units          As Long = 1000

'Read-Only
Private m_int_CurrentPage               As Integer
Private m_int_DevicesCount              As Integer
Private m_str_DeviceDriver              As String
Private m_str_DeviceNames               As String
Private m_str_DevicePorts               As String
Private m_int_DevicePortsCount          As Integer
Private m_lng_DeviceResolutionX         As Long
Private m_lng_DeviceResolutionY         As Long
Private m_enum_Error                    As PrinterErrorConstants
Private m_bool_IsPaperSize              As Boolean
Private m_bool_IsPaperBin               As Boolean
Private m_sng_TwipsPerPixelX            As Single
Private m_sng_TwipsPerPixelY            As Single

'Read/Write
Private m_bool_AbortWindow              As Boolean
Private m_typ_AbortWindowSettings       As ABORTWINDOWTYPE
Private m_enum_Collate                  As CollateConstants
Private m_enum_ColorMode                As ColorModeConstants
Private m_int_Copies                    As Integer
Private m_bool_DefaultDevice            As Boolean
Private m_str_DeviceName                As String
Private m_str_DevicePort                As String
Private m_enum_Duplex                   As DuplexConstants
Private m_str_FileName                  As String
Private m_lng_MarginBottom              As Long
Private m_lng_MarginLeft                As Long
Private m_lng_MarginRight               As Long
Private m_lng_MarginTop                 As Long
Private m_enum_MouseZoom                As MouseZoomConstants
Private m_enum_Orientation              As OrientationConstants
Private m_enum_PageBorder               As PageBorderConstants
Private m_lng_PageHeight                As Long
Private m_lng_PageWidth                 As Long
Private m_enum_PaperBin                 As PaperBinConstants
Private m_lng_PaperHeight               As Long
Private m_int_PaperShadowOffset         As Integer
Private m_enum_PaperSize                As PaperSizeConstants
Private m_lng_PaperWidth                As Long
Private m_bool_PhysicalPage             As Boolean
Private m_bool_Preview                  As Boolean
Private m_int_PreviewPage               As Integer
Private m_enum_PrintQuality             As PrintQualityConstants
Private m_int_PrintScale                As Integer
Private m_enum_ShowGuides               As ShowGuidesConstants
Private m_enum_TrueType                 As TrueTypeFontsConstants
Private m_sng_Zoom                      As Single
Private m_int_ZoomMax                   As Integer
Private m_int_ZoomMin                   As Integer
Private m_enum_ZoomMode                 As ZoomModeConstants
Private m_int_ZoomStep                  As Integer

'Misc
Private m_lng_FromPage                  As Long
Private m_lng_ToPage                    As Long
Private m_int_Width                     As Integer
Private m_int_Height                    As Integer

Private m_typ_PrintDlg                  As TPRINTDLG
Private m_typ_PrintDlgFlags             As PRINTDLGFLAGS
Private m_typ_PageSetupDlg              As TPAGESETUPDLG
Private m_typ_PageSetupDlgFlags            As PAGESETUPDLGFLAGS
Private m_typ_DevMode                   As DEVMODE
Private m_typ_DevNames                  As DEVNAMES

Private m_bool_AuthorMode               As Boolean

'*************************************************************
' CLASS CONTRUCTOR/DECONSTRUCTOR
'*************************************************************
Private Sub UserControl_Initialize()
    #If ShowDebugPrints = 1 Then
        Debug.Print "UserControl_Initialize"
    #End If

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    #If RunStackLogger = 1 Then
        Call LogStackItem("UserControl_Initialize")
    #End If

    'Initiate Scroll Bars
    Set ScrollBars = New cScrollBars
    Call ScrollBars.Create(picWorkspace.hWnd)

    'Initialize Members
    Set Pages = New Pages
    Set Header = New PageHeaderFooter
    Set Footer = New PageHeaderFooter

    'Initiate Printer Structs
    Call vbInitPageSetupDlg(UserControl.hWnd, m_typ_PageSetupDlg, m_typ_PrintDlg, m_typ_DevMode, m_typ_DevNames, puThousandthsOfInches)

    'Initiate First Page
    m_enum_CurrentZoomScale = zsWholePage
    Call pInitPageParams
    Call NewPage

    'Evaluate Printer Device and Display Document
    If Printers.Count <> 0 Then
        m_lng_PaperWidth = Printer.Width
        m_lng_PaperHeight = Printer.Height
        picPage.MouseIcon = LoadResPicture(101, vbResCursor)
        picPage.MousePointer = vbCustom
    Else
        picPage.Visible = False
        picPageShadow.Visible = False
        imgNoPrinters.Visible = True
        lblNoPrinters.Visible = True
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - UserControl_Initialize", Err.Description)
End Sub

Private Sub UserControl_Terminate()
    #If RunStackLogger = 1 Then
        Call LogStackItem("UserControl_Terminate")
    #End If

    Set ScrollBars = Nothing
    Set Pages = Nothing
    Set Header = Nothing
    Set Footer = Nothing
End Sub
'*************************************************************

'*************************************************************
' PUBLIC READ ONLY PROPERTIES
'*************************************************************
Property Get CurrentPage() As Integer
Attribute CurrentPage.VB_Description = "Returns the number of the page being printed."
Attribute CurrentPage.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get CurrentPage")
    #End If
    CurrentPage = m_int_CurrentPage
End Property

Property Get DevicesCount() As Integer
Attribute DevicesCount.VB_Description = "Returns the number of printing devices available."
Attribute DevicesCount.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get DevicesCount")
    #End If
    DevicesCount = Printers.Count    'm_int_DevicesCount
End Property

Property Get DeviceDriver() As String
Attribute DeviceDriver.VB_Description = "Returns the name of the current printer driver."
Attribute DeviceDriver.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get DeviceDriver")
    #End If
    DeviceDriver = m_str_DeviceDriver
End Property

Property Get DeviceNames(ByVal x As Integer) As String
Attribute DeviceNames.VB_Description = "Returns the names of the printing devices available."
Attribute DeviceNames.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get DeviceNames")
    #End If
    DeviceNames = Printers(x).DeviceName    'm_str_DeviceNames
End Property

Property Get DevicePorts(ByVal x As Integer) As String
Attribute DevicePorts.VB_Description = "Returns the names of the ports to which the current printer is connected."
Attribute DevicePorts.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get DevicePorts")
    #End If
    DevicePorts = Printers(x).Port    ' m_str_DevicePorts
End Property

Property Get DevicePortsCount() As Integer
Attribute DevicePortsCount.VB_Description = "Returns the number of ports to which the current printer is connected."
Attribute DevicePortsCount.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get DevicePortsCount")
    #End If
    DevicePortsCount = m_int_DevicePortsCount
End Property

Property Get DeviceResolutionX() As Integer
Attribute DeviceResolutionX.VB_Description = "Returns the number of twips per logical inch along the screen width."
Attribute DeviceResolutionX.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get DeviceResolutionX")
    #End If
    DeviceResolutionX = m_lng_DeviceResolutionX
End Property

Property Get DeviceResolutionY() As Integer
Attribute DeviceResolutionY.VB_Description = "Returns the number of twips per logical inch along the screen height."
Attribute DeviceResolutionY.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get DeviceResolutionY")
    #End If
    DeviceResolutionY = m_lng_DeviceResolutionY
End Property

Property Get Error() As PrinterErrorConstants
Attribute Error.VB_Description = "Returns a code that describes an error condition."
Attribute Error.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get Error")
    #End If

    Error = m_enum_Error
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns the control's current hDC."
Attribute hdc.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get hdc")
    #End If

    hdc = UserControl.hdc
End Property

Property Get IsPaperSize(ByVal x As PaperSizeConstants) As Boolean
Attribute IsPaperSize.VB_Description = "Returns whether a given page size is available on the current printer."
Attribute IsPaperSize.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get IsPaperSize")
    #End If

    On Error Resume Next

    Printer.PaperSize = x
    If Err Then
        Call Err.Clear
        IsPaperSize = False
    Else
        IsPaperSize = True
    End If
    'IsPaperSize = m_bool_IsPaperSize
End Property

Property Get IsPaperBin(ByVal x As PaperBinConstants) As Boolean
Attribute IsPaperBin.VB_Description = "Returns whether a given paper bin is available on the current printer."
Attribute IsPaperBin.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get IsPaperBin")
    #End If

    On Error Resume Next

    Printer.PaperBin = x

    If Err Then
        Call Err.Clear
        IsPaperBin = False
    Else
        IsPaperBin = True
    End If
    'IsPaperBin = m_bool_IsPaperBin
End Property

Property Get TwipsPerPixelX() As Single
Attribute TwipsPerPixelX.VB_Description = "Returns the number of twips per printer pixel in the horizontal direction."
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get TwipsPerPixelX")
    #End If

    TwipsPerPixelX = m_sng_TwipsPerPixelX
End Property

Property Get TwipsPerPixelY() As Single
Attribute TwipsPerPixelY.VB_Description = "Returns the number of twips per printer pixel in the vertical direction."
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get TwipsPerPixelY")
    #End If

    TwipsPerPixelY = m_sng_TwipsPerPixelY
End Property

Property Get Version() As Single
Attribute Version.VB_Description = "Returns the version of the control currently loaded."
Attribute Version.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get Version")
    #End If

    Version = CSng((App.Major & "." & App.Minor))
End Property
'*************************************************************

'*************************************************************
' PUBLIC READ/WRITE PROPERTIES
'*************************************************************
Property Get AbortWindow() As Boolean
Attribute AbortWindow.VB_Description = "Returns or sets whether an Abort dialog will appear while the control is printing."
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get AbortWindow")
    #End If

    AbortWindow = m_bool_AbortWindow
End Property
Property Let AbortWindow(x As Boolean)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let AbortWindow")
    #End If

    m_bool_AbortWindow = x
    Call PropertyChanged("AbortWindow")
End Property

Property Get AbortWindowSettings() As ABORTWINDOWTYPE
Attribute AbortWindowSettings.VB_Description = "Returns or sets the settings for the default Abort dialog."
Attribute AbortWindowSettings.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get AbortWindowSettings")
    #End If

    AbortWindowSettings = m_typ_AbortWindowSettings
End Property
Property Let AbortWindowSettings(x As ABORTWINDOWTYPE)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let AbortWindowSettings")
    #End If

    m_typ_AbortWindowSettings = x
    Call PropertyChanged("AbortWindowSettings")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get Appearance")
    #End If

    Appearance = UserControl.Appearance
End Property
Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let Appearance")
    #End If

    UserControl.Appearance() = New_Appearance
    Call PropertyChanged("Appearance")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picWorkspace,picWorkspace,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns or sets the color of the workspace around the preview page."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get BackColor")
    #End If

    BackColor = picWorkspace.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let BackColor")
    #End If

    picWorkspace.BackColor() = New_BackColor
    Call PropertyChanged("BackColor")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get BorderStyle")
    #End If

    BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let BorderStyle")
    #End If

    UserControl.BorderStyle() = New_BorderStyle
    Call PropertyChanged("BorderStyle")
End Property

Property Get Collate() As CollateConstants
Attribute Collate.VB_Description = "Returns or sets whether multiple copies will be collated."
Attribute Collate.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get Collate")
    #End If

    Collate = m_enum_Collate
End Property
Property Let Collate(x As CollateConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let Collate")
    #End If

    m_enum_Collate = x
    m_typ_DevMode.dmCollate = x
    Call PropertyChanged("Collate")
End Property

Property Get ColorMode() As ColorModeConstants
Attribute ColorMode.VB_Description = "Returns or sets the color mode on color printers."
Attribute ColorMode.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get ColorMode")
    #End If

    ColorMode = m_enum_ColorMode
End Property
Property Let ColorMode(x As ColorModeConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let ColorMode")
    #End If

    m_enum_ColorMode = x
    m_typ_DevMode.dmColor = x
    Call PropertyChanged("ColorMode")
End Property

Property Get Copies() As Integer
Attribute Copies.VB_Description = "Returns/sets a value that determines the number of copies to be printed."
Attribute Copies.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get Copies")
    #End If

    Copies = m_int_Copies
End Property
Property Let Copies(x As Integer)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let Copies")
    #End If

    m_int_Copies = x
    m_typ_DevMode.dmCopies = x
    Call PropertyChanged("Copies")
End Property

Property Get DefaultDevice() As Boolean
Attribute DefaultDevice.VB_Description = "Returns or sets whether device changes affect the Windows default settings."
Attribute DefaultDevice.VB_ProcData.VB_Invoke_Property = ";Behavior"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get DefaultDevice")
    #End If

    DefaultDevice = m_bool_DefaultDevice
End Property
Property Let DefaultDevice(x As Boolean)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let DefaultDevice")
    #End If

    m_bool_DefaultDevice = x
    Call PropertyChanged("DefaultDevice")
End Property

Property Get DeviceName() As String
Attribute DeviceName.VB_Description = "Returns the name of the device a driver supports."
Attribute DeviceName.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get DeviceName")
    #End If

    DeviceName = Printer.DeviceName
    'DeviceName = m_str_DeviceName
End Property
Property Let DeviceName(x As String)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let DeviceName")
    #End If

    m_str_DeviceName = x
    Call PropertyChanged("DeviceName")
End Property

Property Get DevicePort() As String
Attribute DevicePort.VB_Description = "Returns the name of the port through which a document is sent to a printer."
Attribute DevicePort.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get DevicePort")
    #End If

    DevicePort = Printer.Port
    'DevicePort = m_str_DevicePort
End Property
Property Let DevicePort(x As String)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let DevicePort")
    #End If

    m_str_DevicePort = x
    Call PropertyChanged("DevicePort")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPage,picPage,-1,DrawStyle
Public Property Get DrawStyle() As DrawStyleConstants
Attribute DrawStyle.VB_Description = "Determines the line style for output from graphics methods."
Attribute DrawStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get DrawStyle")
    #End If

    DrawStyle = picPage.DrawStyle
End Property
Public Property Let DrawStyle(ByVal New_DrawStyle As DrawStyleConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let DrawStyle")
    #End If

    picPage.DrawStyle() = New_DrawStyle
    Call PropertyChanged("DrawStyle")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPage,picPage,-1,DrawWidth
Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "Returns/sets the line width for output from graphics methods."
Attribute DrawWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get DrawWidth")
    #End If

    DrawWidth = picPage.DrawWidth
End Property
Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let DrawWidth")
    #End If

    picPage.DrawWidth() = New_DrawWidth
    Call PropertyChanged("DrawWidth")
End Property

Property Get Duplex() As DuplexConstants
Attribute Duplex.VB_Description = "Determines whether a page is printed on both sides."
Attribute Duplex.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get Duplex")
    #End If

    Duplex = m_enum_Duplex
End Property
Property Let Duplex(x As DuplexConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let Duplex")
    #End If

    m_enum_Duplex = x
    m_typ_DevMode.dmDuplex = x
    Call PropertyChanged("Duplex")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get Enabled")
    #End If

    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let Enabled")
    #End If

    UserControl.Enabled() = New_Enabled
    Call PropertyChanged("Enabled")
End Property

Property Get FileName() As String
Attribute FileName.VB_Description = "Returns or sets the name of the current document."
Attribute FileName.VB_ProcData.VB_Invoke_Property = ";Misc"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get FileName")
    #End If

    FileName = m_str_FileName
End Property
Property Let FileName(x As String)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let FileName")
    #End If

    m_str_FileName = x
    Call PropertyChanged("FileName")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPage,picPage,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in closed shapes like circles, boxes, etc."
Attribute FillColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get FillColor")
    #End If

    FillColor = picPage.FillColor
End Property
Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let FillColor")
    #End If

    picPage.FillColor() = New_FillColor
    Call PropertyChanged("FillColor")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPage,picPage,-1,FillStyle
Public Property Get FillStyle() As FillStyleConstants
Attribute FillStyle.VB_Description = "Returns/sets the fill style of a shape."
Attribute FillStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get FillStyle")
    #End If

    FillStyle = picPage.FillStyle
End Property
Public Property Let FillStyle(ByVal New_FillStyle As FillStyleConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let FillStyle")
    #End If

    picPage.FillStyle() = New_FillStyle
    Call PropertyChanged("FillStyle")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPage,picPage,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get Font")
    #End If

    Set Font = picPage.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let Font")
    #End If

    Set picPage.Font = New_Font
    Call PropertyChanged("Font")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPage,picPage,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
Attribute FontBold.VB_ProcData.VB_Invoke_Property = ";Font"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get FontBold")
    #End If

    FontBold = picPage.FontBold
End Property
Public Property Let FontBold(ByVal New_FontBold As Boolean)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let FontBold")
    #End If

    picPage.FontBold() = New_FontBold
    Call PropertyChanged("FontBold")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPage,picPage,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
Attribute FontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get FontItalic")
    #End If

    FontItalic = picPage.FontItalic
End Property
Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let FontItalic")
    #End If

    picPage.FontItalic() = New_FontItalic
    Call PropertyChanged("FontItalic")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPage,picPage,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
Attribute FontName.VB_ProcData.VB_Invoke_Property = ";Font"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get FontName")
    #End If

    FontName = picPage.FontName
End Property
Public Property Let FontName(ByVal New_FontName As String)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let FontName")
    #End If

    picPage.FontName() = New_FontName
    Call PropertyChanged("FontName")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPage,picPage,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
Attribute FontSize.VB_ProcData.VB_Invoke_Property = ";Font"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get FontSize")
    #End If

    FontSize = picPage.FontSize
End Property
Public Property Let FontSize(ByVal New_FontSize As Single)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let FontSize")
    #End If

    picPage.FontSize() = New_FontSize
    Call PropertyChanged("FontSize")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPage,picPage,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
Attribute FontStrikethru.VB_ProcData.VB_Invoke_Property = ";Font"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get FontStrikethru")
    #End If

    FontStrikethru = picPage.FontStrikethru
End Property
Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let FontStrikethru")
    #End If

    picPage.FontStrikethru() = New_FontStrikethru
    Call PropertyChanged("FontStrikethru")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPage,picPage,-1,FontTransparent
Public Property Get FontTransparent() As Boolean
Attribute FontTransparent.VB_Description = "Returns/sets a value that determines whether background text/graphics on a Form, Printer or PictureBox are displayed."
Attribute FontTransparent.VB_ProcData.VB_Invoke_Property = ";Font"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get FontTransparent")
    #End If

    FontTransparent = picPage.FontTransparent
End Property
Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let FontTransparent")
    #End If

    picPage.FontTransparent() = New_FontTransparent
    Call PropertyChanged("FontTransparent")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPage,picPage,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
Attribute FontUnderline.VB_ProcData.VB_Invoke_Property = ";Font"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get FontUnderline")
    #End If

    FontUnderline = picPage.FontUnderline
End Property
Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let FontUnderline")
    #End If

    picPage.FontUnderline() = New_FontUnderline
    Call PropertyChanged("FontUnderline")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPage,picPage,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get ForeColor")
    #End If

    ForeColor = picPage.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let ForeColor")
    #End If

    picPage.ForeColor() = New_ForeColor
    Call PropertyChanged("ForeColor")
End Property

Property Get MarginBottom() As Long
Attribute MarginBottom.VB_Description = "Returns or sets the bottom margin, in twips."
Attribute MarginBottom.VB_ProcData.VB_Invoke_Property = ";Text"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get MarginBottom")
    #End If

    MarginBottom = m_lng_MarginBottom
End Property
Property Let MarginBottom(x As Long)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let MarginBottom")
    #End If

    m_lng_MarginBottom = x
    m_typ_PageSetupDlg.rtMargin.bottom = pConvertTwipsToPrinterUnits(x)
    Call pUpdatePageMargins
    Call PropertyChanged("MarginBottom")
End Property

Property Get MarginLeft() As Long
Attribute MarginLeft.VB_Description = "Returns or sets the left margin, in twips."
Attribute MarginLeft.VB_ProcData.VB_Invoke_Property = ";Text"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get MarginLeft")
    #End If

    MarginLeft = m_lng_MarginLeft
End Property
Property Let MarginLeft(x As Long)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let MarginLeft")
    #End If

    m_lng_MarginLeft = x
    m_typ_PageSetupDlg.rtMargin.Left = pConvertTwipsToPrinterUnits(x)
    Call pUpdatePageMargins
    Call PropertyChanged("MarginLeft")
End Property

Property Get MarginRight() As Long
Attribute MarginRight.VB_Description = "Returns or sets the right margin, in twips."
Attribute MarginRight.VB_ProcData.VB_Invoke_Property = ";Text"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get MarginRight")
    #End If

    MarginRight = m_lng_MarginRight
End Property
Property Let MarginRight(x As Long)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let MarginRight")
    #End If

    m_lng_MarginRight = x
    m_typ_PageSetupDlg.rtMargin.right = pConvertTwipsToPrinterUnits(x)
    Call pUpdatePageMargins
    Call PropertyChanged("MarginRight")
End Property

Property Get MarginTop() As Long
Attribute MarginTop.VB_Description = "Returns or sets the top margin, in twips."
Attribute MarginTop.VB_ProcData.VB_Invoke_Property = ";Text"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get MarginTop")
    #End If

    MarginTop = m_lng_MarginTop
End Property
Property Let MarginTop(x As Long)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let MarginTop")
    #End If

    m_lng_MarginTop = x
    m_typ_PageSetupDlg.rtMargin.Top = pConvertTwipsToPrinterUnits(x)
    Call pUpdatePageMargins
    Call PropertyChanged("MarginTop")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Misc"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get MouseIcon")
    #End If

    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Set MouseIcon")
    #End If

    Set UserControl.MouseIcon = New_MouseIcon
    Call PropertyChanged("MouseIcon")
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Misc"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get MousePointer")
    #End If

    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let MousePointer")
    #End If

    UserControl.MousePointer() = New_MousePointer
    Call PropertyChanged("MousePointer")
End Property

Property Get MouseZoom() As MouseZoomConstants
Attribute MouseZoom.VB_Description = "Returns or sets whether the user can zoom in and out by double clicking on the preview window."
Attribute MouseZoom.VB_ProcData.VB_Invoke_Property = ";Misc"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get MouseZoom")
    #End If

    MouseZoom = m_enum_MouseZoom
End Property
Property Let MouseZoom(x As MouseZoomConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let MouseZoom")
    #End If

    m_enum_MouseZoom = x
    Call PropertyChanged("MouseZoom")
End Property

Property Get Orientation() As OrientationConstants
Attribute Orientation.VB_Description = "Returns or sets the paper orientation."
Attribute Orientation.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get Orientation")
    #End If

    Orientation = m_enum_Orientation
End Property
Property Let Orientation(x As OrientationConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let Orientation")
    #End If

    m_enum_Orientation = x
    m_typ_DevMode.dmOrientation = x
    Call pUpdatePageMargins
    Call PropertyChanged("Orientation")
End Property

Property Get PageBorder() As PageBorderConstants
Attribute PageBorder.VB_Description = "Returns or sets the type of border to draw around each page."
Attribute PageBorder.VB_ProcData.VB_Invoke_Property = ";Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PageBorder")
    #End If

    PageBorder = m_enum_PageBorder
End Property
Property Let PageBorder(x As PageBorderConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PageBorder")
    #End If

    m_enum_PageBorder = x
    Call PropertyChanged("PageBorder")
End Property

Property Get PageHeight() As Long
Attribute PageHeight.VB_Description = "Returns or sets the height of the printable area on the page, in twips"
Attribute PageHeight.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PageHeight")
    #End If

    PageHeight = m_lng_PageHeight
End Property
Property Let PageHeight(x As Long)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PageHeight")
    #End If

    m_lng_PageHeight = x
    Call PropertyChanged("PageHeight")
End Property

Property Get PageWidth() As Long
Attribute PageWidth.VB_Description = "Returns or sets the width of the printable area on the page, in twips"
Attribute PageWidth.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PageWidth")
    #End If

    PageWidth = m_lng_PageWidth
End Property
Property Let PageWidth(x As Long)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PageWidth")
    #End If

    m_lng_PageWidth = x
    Call PropertyChanged("PageWidth")
End Property

Property Get PaperBin() As PaperBinConstants
Attribute PaperBin.VB_Description = "Returns or sets the paper bin to use."
Attribute PaperBin.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PaperBin")
    #End If

    PaperBin = m_enum_PaperBin
End Property
Property Let PaperBin(x As PaperBinConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PaperBin")
    #End If

    Printer.PaperBin = x
    m_enum_PaperBin = Printer.PaperBin
    Call PropertyChanged("PaperBin")
End Property

Property Get PaperBorderColor() As OLE_COLOR
Attribute PaperBorderColor.VB_Description = "Returns or sets the color used for the paper border"
Attribute PaperBorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PaperBorderColor")
    #End If

    PaperBorderColor = shpPaperBorder.BorderColor
End Property
Property Let PaperBorderColor(x As OLE_COLOR)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PaperBorderColor")
    #End If

    shpPaperBorder.BorderColor = x
    Call PropertyChanged("PaperBorderColor")
End Property

Property Get PaperBorderWidth() As Integer
Attribute PaperBorderWidth.VB_Description = "Returns or sets the line width for the paper border."
Attribute PaperBorderWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PaperBorderWidth")
    #End If

    PaperBorderWidth = shpPaperBorder.BorderWidth
End Property
Property Let PaperBorderWidth(x As Integer)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PaperBorderWidth")
    #End If

    If x <= 0 Then Exit Property
    shpPaperBorder.BorderWidth = x
    Call PropertyChanged("PaperBorderWidth")
End Property

Property Get PaperHeight() As Long
Attribute PaperHeight.VB_Description = "Returns or sets the height of a custom paper size, in twips."
Attribute PaperHeight.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PaperHeight")
    #End If

    PaperHeight = m_lng_PaperHeight
End Property
Property Let PaperHeight(x As Long)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PaperHeight")
    #End If

    m_lng_PaperHeight = x
    m_typ_PageSetupDlg.ptPaperSize.Y = pConvertTwipsToPrinterUnits(x)
    Call PropertyChanged("PaperHeight")
End Property

Property Get PaperShadowBorderStyle() As BorderStyleConstants
Attribute PaperShadowBorderStyle.VB_Description = "Returns or sets the line style for border of the Paper Shadow."
Attribute PaperShadowBorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PaperShadowBorderStyle")
    #End If

    PaperShadowBorderStyle = picPageShadow.BorderStyle
End Property
Property Let PaperShadowBorderStyle(x As BorderStyleConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PaperShadowBorderStyle")
    #End If

    picPageShadow.BorderStyle = x
    Call PropertyChanged("PaperShadowBorderStyle")
End Property

Property Get PaperShadowColor() As OLE_COLOR
Attribute PaperShadowColor.VB_Description = "Returns or sets the color of the Paper Shadow."
Attribute PaperShadowColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PaperShadowColor")
    #End If

    PaperShadowColor = picPageShadow.BackColor
End Property
Property Let PaperShadowColor(x As OLE_COLOR)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PaperShadowColor")
    #End If

    picPageShadow.BackColor = x
    Call PropertyChanged("PaperShadowColor")
End Property

Property Get PaperShadowOffset() As Integer
Attribute PaperShadowOffset.VB_Description = "Returns or sets the distance (twips) the shadow if offset from the paper in the X and Y-axis."
Attribute PaperShadowOffset.VB_ProcData.VB_Invoke_Property = ";Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PaperShadowOffset")
    #End If

    PaperShadowOffset = m_int_PaperShadowOffset
End Property
Property Let PaperShadowOffset(x As Integer)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PaperShadowOffset")
    #End If

    m_int_PaperShadowOffset = x
    Call picWorkspace_Resize
    Call PropertyChanged("PaperShadowOffset")
End Property

Property Get PaperSize() As PaperSizeConstants
Attribute PaperSize.VB_Description = "Returns or sets a standard paper size."
Attribute PaperSize.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PaperSize")
    #End If

    PaperSize = m_enum_PaperSize
End Property
Property Let PaperSize(x As PaperSizeConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PaperSize")
    #End If

    m_enum_PaperSize = x
    m_typ_DevMode.dmPaperSize = x

    Select Case x
        Case psLetter    'Letter, 8 1/2 x 11 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(8.5)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(11)
        Case psLetterSmall    'Letter Small, 8 1/2 x 11 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(8.5)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(11)
        Case psTabloid    'Tabloid, 11 x 17 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(11)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(17)
        Case psLedger    'Ledger, 17 x 11 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(17)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(11)
        Case psLegal    'Legal, 8 1/2 x 14 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(8.5)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(14)
        Case psStatement    'Statement, 5 1/2 x 8 1/2 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(5.5)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(8.5)
        Case psExecutive    'Executive, 7 1/2 x 10 1/2 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(7.5)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(10.5)
        Case psA3    'A3, 297 x 420 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(297, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(420, vbMillimeters, vbInches))
        Case psA4    'A4, 210 x 297 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(210, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(297, vbMillimeters, vbInches))
        Case psA4Small    'A4 Small, 210 x 297 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(210, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(297, vbMillimeters, vbInches))
        Case psA5    'A5, 148 x 210 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(148, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(210, vbMillimeters, vbInches))
        Case psB4    'B4, 250 x 354 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(250, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(354, vbMillimeters, vbInches))
        Case psB5    'B5, 182 x 257 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(182, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(257, vbMillimeters, vbInches))
        Case psFolio    'Folio, 8 1/2 x 13 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(8.5)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(13)
        Case psQuarto    'Quarto, 215 x 275 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(215, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(275, vbMillimeters, vbInches))
        Case ps10x14    '10 x 14 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(10)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(14)
        Case ps11x17    '11 x 17 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(11)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(17)
        Case psNote    'Note, 8 1/2 x 11 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(8.5)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(11)
        Case psEnv9    'Envelope #9, 3 7/8 x 8 7/8 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(3.875)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(8.875)
        Case psEnv10    'Envelope #10, 4 1/8 x 9 1/2 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(4.125)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(9.5)
        Case psEnv11    'Envelope #11, 4 1/2 x 10 3/8 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(4.5)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(10.375)
        Case psEnv12    'Envelope #12, 4 1/2 x 11 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(4.5)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(11)
        Case psEnv14    'Envelope #14, 5 x 11 1/2 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(5)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(11.5)
        Case psCSheet    'C size sheet
        Case psDSheet    'D size sheet
        Case psESheet    'E size sheet
        Case psEnvDL    'Envelope DL, 110 x 220 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(110, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(220, vbMillimeters, vbInches))
        Case psEnvC3    'Envelope C3, 324 x 458 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(324, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(458, vbMillimeters, vbInches))
        Case psEnvC4    'Envelope C4, 229 x 324 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(229, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(324, vbMillimeters, vbInches))
        Case psEnvC5    'Envelope C5, 162 x 229 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(162, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(229, vbMillimeters, vbInches))
        Case psEnvC6    'Envelope C6, 114 x 162 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(114, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(162, vbMillimeters, vbInches))
        Case psEnvC65    'Envelope C65, 114 x 229 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(114, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(229, vbMillimeters, vbInches))
        Case psEnvB4    'Envelope B4, 250 x 353 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(250, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(353, vbMillimeters, vbInches))
        Case psEnvB5    'Envelope B5, 176 x 250 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(176, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(250, vbMillimeters, vbInches))
        Case psEnvB6    'Envelope B6, 176 x 125 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(176, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(125, vbMillimeters, vbInches))
        Case psEnvItaly    'Envelope, 110 x 230 mm
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(ScaleX(110, vbMillimeters, vbInches))
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(ScaleY(230, vbMillimeters, vbInches))
        Case psEnvMonarch    'Envelope Monarch, 3 7/8 x 7 1/2 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(8.875)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(7.5)
        Case psEnvPersonal    'Envelope, 3 5/8 x 6 1/2 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(3.625)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(6.5)
        Case psFanfoldUS    'U.S. Standard Fanfold, 14 7/8 x 11 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(14.875)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(11)
        Case psFanfoldStdGerman    'German Standard Fanfold, 8 1/2 x 12 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(8.5)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(12)
        Case psFanfoldLglGerman    'German Legal Fanfold, 8 1/2 x 13 in.
            m_typ_PageSetupDlg.ptPaperSize.x = pConvertInchesToPrinterUnits(8.5)
            m_typ_PageSetupDlg.ptPaperSize.Y = pConvertInchesToPrinterUnits(13)
        Case psUser    'User-defined
    End Select

    Call pUpdatePageMargins
    Call PropertyChanged("PaperSize")
End Property

Property Get PaperWidth() As Long
Attribute PaperWidth.VB_Description = "Returns or sets the width of a custom paper size, in twips."
Attribute PaperWidth.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PaperWidth")
    #End If

    PaperWidth = m_lng_PaperWidth
End Property
Property Let PaperWidth(x As Long)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PaperWidth")
    #End If

    m_lng_PaperWidth = x
    m_typ_PageSetupDlg.ptPaperSize.x = pConvertTwipsToPrinterUnits(x)
    Call picWorkspace_Resize
    Call PropertyChanged("PaperWidth")
End Property

Property Get PhysicalPage() As Boolean
Attribute PhysicalPage.VB_Description = "Returns or sets whether to use the physical size of the page or on its printable area."
Attribute PhysicalPage.VB_ProcData.VB_Invoke_Property = ";Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PhysicalPage")
    #End If

    PhysicalPage = m_bool_PhysicalPage
End Property
Property Let PhysicalPage(x As Boolean)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PhysicalPage")
    #End If

    m_bool_PhysicalPage = x
    Call PropertyChanged("PhysicalPage")
End Property

Property Get Preview() As Boolean
Attribute Preview.VB_Description = "Returns or sets whether output saved for previewing or sent directly to the printer."
Attribute Preview.VB_ProcData.VB_Invoke_Property = ";Behavior"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get Preview")
    #End If

    Preview = m_bool_Preview
End Property
Property Let Preview(x As Boolean)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let Preview")
    #End If

    m_bool_Preview = x
    Call PropertyChanged("Preview")
End Property

Property Get PreviewPage() As Integer
Attribute PreviewPage.VB_Description = "Returns or sets the current preview page (first page is 1)."
Attribute PreviewPage.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PreviewPage")
    #End If

    PreviewPage = m_int_PreviewPage
End Property
Property Let PreviewPage(x As Integer)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PreviewPage")
    #End If

    If x > 0 And x <= Pages.Count Then
        m_int_PreviewPage = x
        m_int_CurrentPage = x
        Call picWorkspace_Resize
        Call picPage_Paint
        Call PropertyChanged("PreviewPage")
    Else
        Call Err.Raise(vbObjectError, App.Title, "Invalid Page assignment.")
    End If
End Property

Property Get PrintQuality() As PrintQualityConstants
Attribute PrintQuality.VB_Description = "Returns or sets the print quality."
Attribute PrintQuality.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PrintQuality")
    #End If

    PrintQuality = m_enum_PrintQuality
End Property
Property Let PrintQuality(x As PrintQualityConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PrintQuality")
    #End If

    Printer.PrintQuality = x
    m_typ_DevMode.dmPrintQuality = x
    m_enum_PrintQuality = Printer.PrintQuality
    Call PropertyChanged("PrintQuality")
End Property

Property Get PrintScale() As Integer
Attribute PrintScale.VB_Description = "Returns or sets the percentage by which the printed output is to be scaled."
Attribute PrintScale.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get PrintScale")
    #End If

    PrintScale = m_int_PrintScale
End Property
Property Let PrintScale(x As Integer)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let PrintScale")
    #End If

    m_int_PrintScale = x
    m_typ_DevMode.dmScale = x
    Call PropertyChanged("PrintScale")
End Property

Property Get ShowGuides() As ShowGuidesConstants
Attribute ShowGuides.VB_Description = "Returns or sets whether margin guides are displayed on the page."
Attribute ShowGuides.VB_ProcData.VB_Invoke_Property = ";Appearance"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get ShowGuides")
    #End If

    ShowGuides = m_enum_ShowGuides
End Property
Property Let ShowGuides(x As ShowGuidesConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let ShowGuides")
    #End If

    m_enum_ShowGuides = x
    Call PropertyChanged("ShowGuides")
End Property

Property Get TrueType() As TrueTypeFontsConstants
Attribute TrueType.VB_Description = "Returns or sets how TrueType fonts should be printed."
Attribute TrueType.VB_MemberFlags = "400"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get TrueType")
    #End If

    TrueType = m_enum_TrueType
End Property
Property Let TrueType(x As TrueTypeFontsConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let TrueType")
    #End If

    m_enum_TrueType = x
    m_typ_DevMode.dmTTOption = x
    Call PropertyChanged("TrueType")
End Property

Property Get Zoom() As Single
Attribute Zoom.VB_Description = "Returns or sets the preview scale: set to a percentage, or zero to fill the control."
Attribute Zoom.VB_ProcData.VB_Invoke_Property = ";Behavior"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get Zoom")
    #End If

    Zoom = m_sng_Zoom
End Property
Property Let Zoom(x As Single)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let Zoom")
    #End If

    m_sng_Zoom = x
    Call PropertyChanged("Zoom")
End Property

Property Get ZoomMax() As Integer
Attribute ZoomMax.VB_Description = "Returns or sets the maximum valid zoom factor."
Attribute ZoomMax.VB_ProcData.VB_Invoke_Property = ";Behavior"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get ZoomMax")
    #End If

    ZoomMax = m_int_ZoomMax
End Property
Property Let ZoomMax(x As Integer)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let ZoomMax")
    #End If

    m_int_ZoomMax = x
    Call PropertyChanged("ZoomMax")
End Property

Property Get ZoomMin() As Integer
Attribute ZoomMin.VB_Description = "Returns or sets the minimum valid zoom factor."
Attribute ZoomMin.VB_ProcData.VB_Invoke_Property = ";Behavior"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get ZoomMin")
    #End If

    ZoomMin = m_int_ZoomMin
End Property
Property Let ZoomMin(x As Integer)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let ZoomMin")
    #End If

    m_int_ZoomMin = x
    Call PropertyChanged("ZoomMin")
End Property

Property Get ZoomMode() As ZoomModeConstants
Attribute ZoomMode.VB_Description = "Sets or returns the zoom mode."
Attribute ZoomMode.VB_ProcData.VB_Invoke_Property = ";Behavior"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get ZoomMode")
    #End If

    ZoomMode = m_enum_ZoomMode
End Property
Property Let ZoomMode(x As ZoomModeConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let ZoomMode")
    #End If

    m_enum_ZoomMode = x
    Call PropertyChanged("ZoomMode")
End Property

Property Get ZoomStep() As Integer
Attribute ZoomStep.VB_Description = "Returns or sets the step for zooming with the mouse."
Attribute ZoomStep.VB_ProcData.VB_Invoke_Property = ";Behavior"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Get ZoomStep")
    #End If

    ZoomStep = m_int_ZoomStep
End Property
Property Let ZoomStep(x As Integer)
    #If RunStackLogger = 1 Then
        Call LogStackItem("Let ZoomStep")
    #End If

    m_int_ZoomStep = x
    Call PropertyChanged("ZoomStep")
End Property
'*************************************************************

'*************************************************************
' PUBLIC METHODS
'*************************************************************
Public Sub Action(ActionType As ActionConstants)
Attribute Action.VB_Description = "Executes an action such as 'StartDoc' or 'EndDoc'"
    #If RunStackLogger = 1 Then
        Call LogStackItem("Action")
    #End If

    '
End Sub

Public Sub ClientToPage(ByVal x As Single, ByVal Y As Single, Optional ByVal Page As Variant)
Attribute ClientToPage.VB_Description = "Converts mouse coordinates to page coordinates."
    #If RunStackLogger = 1 Then
        Call LogStackItem("ClientToPage")
    #End If

    '
End Sub

Public Sub EndDoc()
Attribute EndDoc.VB_Description = "Ends the current document, started with the 'StartDoc' method."
    #If RunStackLogger = 1 Then
        Call LogStackItem("EndDoc")
    #End If

    m_bool_AuthorMode = False
End Sub

Public Sub EndOverlay()
Attribute EndOverlay.VB_Description = "Closes a page that was opened with the StartOverlay method."
    #If RunStackLogger = 1 Then
        Call LogStackItem("EndOverlay")
    #End If

    '
End Sub

Public Sub KillDoc()
Attribute KillDoc.VB_Description = "Cancels and deletes the current document."
    #If RunStackLogger = 1 Then
        Call LogStackItem("KillDoc")
    #End If

    '
End Sub

Public Sub LoadDocument(ByVal FileName As String, Optional ByVal Append As Variant)
Attribute LoadDocument.VB_Description = "Loads a document from disk."
    #If RunStackLogger = 1 Then
        Call LogStackItem("LoadDocument")
    #End If

    '
End Sub

Public Function NewPage() As Page
Attribute NewPage.VB_Description = "Adds a Page object to a Pages collection and returns a reference to the created object."
    #If RunStackLogger = 1 Then
        Call LogStackItem("NewPage")
    #End If

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    Dim PageNew As Page

    Set PageNew = New Page
    Call picPage.Cls

    With PageNew
        .PrinterObject = picPage
        .DrawStyle = picPage.DrawStyle
        .DrawWidth = picPage.DrawWidth
        .FillColor = picPage.FillColor
        .FillColor = picPage.FillColor
        .FillStyle = picPage.FillStyle
        .Font = picPage.Font
        .FontBold = picPage.FontBold
        .FontItalic = picPage.FontItalic
        .FontName = picPage.FontName
        .FontSize = picPage.FontSize
        .FontStrikethru = picPage.FontStrikethru
        .FontTransparent = picPage.FontTransparent
        .FontUnderline = picPage.FontUnderline
        .ForeColor = picPage.ForeColor
        .MarginBottom = m_lng_MarginBottom
        .MarginLeft = m_lng_MarginLeft
        .MarginRight = m_lng_MarginRight
        .MarginTop = m_lng_MarginTop
        .Orientation = m_enum_Orientation
        .PageBorder = m_enum_PageBorder
        .PageNumber = Pages.Count + 1
        .PageWidth = m_lng_PageWidth
        .PageHeight = m_lng_PageHeight
        .PaperHeight = m_lng_PaperHeight
        .PaperSize = m_enum_PaperSize
        .PaperWidth = m_lng_PaperWidth
        .ShowGuides = m_enum_ShowGuides
        Set .Header = Header
        Set .Footer = Footer
    End With

    Call Pages.Add(PageNew)
    Set NewPage = Pages(Pages.Count)
    m_int_CurrentPage = Pages.Count

    Exit Function

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - NewPage", Err.Description)
End Function

Public Sub PageToClient(ByVal x As Single, ByVal Y As Single)
Attribute PageToClient.VB_Description = "Converts page coordinates to mouse coordinates"
    #If RunStackLogger = 1 Then
        Call LogStackItem("PageToClient")
    #End If

    '
End Sub

Public Function PrintDocument(Optional ByVal ShowPrintDlg As Boolean = False, Optional ByVal FromPage As Long, Optional ByVal ToPage As Long) As Boolean
Attribute PrintDocument.VB_Description = "Prints the current document being previewed on the printer."
    #If RunStackLogger = 1 Then
        Call LogStackItem("PrintDocument")
    #End If

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    Dim i               As Integer

    Call pValidateFromToPages(FromPage, ToPage)

    g_bool_SendToPrinter = True

    If ShowPrintDlg Then
        'Set the Flags
        With m_typ_PrintDlgFlags
            .AllPages = True
            .Collate = False
            .DisablePrintToFile = True
            .EnablePrintHook = False
            .EnablePrintTemplate = False
            .EnablePrintTemplateHandle = False
            .EnableSetupHook = False
            .EnableSetupTemplate = False
            .EnableSetupTemplateHandle = False
            .HidePrintToFile = True
            .NoNetworkButton = False
            .NoPageNums = False
            .NoSelection = True
            .NoWarning = False
            .PageNums = False
            .PrintSetup = False
            .PrintToFile = False
            .ReturnDC = False
            .ReturnDefault = False
            .ReturnIC = False
            .Selection = False
            .ShowHelp = False
            .UseDevModeCopies = False
            .UseDevModeCopiesAndCollate = True
        End With

        m_typ_PrintDlg.nFromPage = 1
        m_typ_PrintDlg.nToPage = Pages.Count
        m_typ_PrintDlg.nMinPage = 1
        m_typ_PrintDlg.nMaxPage = Pages.Count    '&HFFFF

        If vbPrintDlg( _
                UserControl.hWnd, _
                m_typ_PrintDlg, _
                m_typ_DevMode, _
                m_typ_DevNames, _
                m_typ_PrintDlgFlags) Then

            Set Printer = Printers(g_int_CurrentPrinterIndex)

            For i = 1 To Pages.Count
                If i <> 1 Then Call PrintNewPage
                If (m_typ_PrintDlg.flags And PD_PAGENUMS) And i >= m_typ_PrintDlg.nFromPage And i <= m_typ_PrintDlg.nToPage Then
                    If PrintByPage(Pages(i), False) = False Then Exit For
                ElseIf (m_typ_PrintDlg.flags And PD_ALLPAGES) Then
                    If PrintByPage(Pages(i), False) = False Then Exit For
                End If
            Next i

            Call PrintEndDoc

            PrintDocument = True
        Else
            PrintDocument = False
        End If
    Else
        For i = 1 To Pages.Count
            If i <> 1 Then Call PrintNewPage
            If i >= FromPage And i <= ToPage Then
                Call PrintByPage(Pages(i), False)
            End If
        Next i

        Call PrintEndDoc
    End If

    g_bool_SendToPrinter = False

    Exit Function

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintDocument", Err.Description)
End Function

Public Sub PrintFile(ByVal FileName As String)
Attribute PrintFile.VB_Description = "Prints a file."
    #If RunStackLogger = 1 Then
        Call LogStackItem("PrintFile")
    #End If

    '
End Sub
'
'Public Sub Refresh()
'    #If RunStackLogger = 1 Then
'        Call LogStackItem("Refresh")
'    #End If
'
'    Call picPage_Paint
'End Sub

Public Sub StartNewDocment()
Attribute StartNewDocment.VB_Description = "Clears the control and starts a new document."
    #If RunStackLogger = 1 Then
        Call LogStackItem("StartNewDocment")
    #End If

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    'Lock the paint messages
    m_bool_AuthorMode = True

    'Initialize Members
    Set Pages = New Pages
    Set Header = New PageHeaderFooter
    Set Footer = New PageHeaderFooter

    Call NewPage

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - StartNewDocment", Err.Description)
End Sub

Public Sub StartOverlay(ByVal Page As Integer, Optional ByVal PreservePrev As Boolean = False)
Attribute StartOverlay.VB_Description = "Reopens an existing preview page for additional output."
    #If RunStackLogger = 1 Then
        Call LogStackItem("StartOverlay")
    #End If

    '
End Sub

Public Sub SaveDocument(ByVal FileName As String, Optional ByVal Compress As Variant, Optional ByVal FromPage As Variant, Optional ByVal ToPage As Variant)
Attribute SaveDocument.VB_Description = "Saves the current document to disk."
    #If RunStackLogger = 1 Then
        Call LogStackItem("SaveDocument")
    #End If

    '
End Sub

Public Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Shows the About Dialog."
    #If RunStackLogger = 1 Then
        Call LogStackItem("ShowAbout")
    #End If

    Call frmAbout.Show(vbModal)
End Sub

Public Function ShowPrintDialog(ByVal DialogType As PrintDialogConstants) As Boolean
    #If RunStackLogger = 1 Then
        Call LogStackItem("ShowPrintDialog - Enter")
    #End If

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    #If ShowDebugPrints = 1 Then
        Debug.Print "***************Checking***************"
        Debug.Print "Printer.ColorMode:", Printer.ColorMode
        Debug.Print "Printer.Copies:", Printer.Copies
        Debug.Print "Printer.Duplex:", Printer.Duplex
        Debug.Print "Printer.Orientation:", Printer.Orientation
        Debug.Print "Printer.PaperBin:", Printer.PaperBin
        Debug.Print "Printer.PaperSize:", Printer.PaperSize
        Debug.Print "Printer.PrintQuality:", Printer.PrintQuality
        Debug.Print "***************DONE***************"
    #End If

    If DialogType = pdPageSetup Then
        With m_typ_PageSetupDlgFlags
            .DefaultMinMargins = False
            .DisableMargins = False
            .DisableOrientation = False
            .DisablePagePainting = False
            .DisablePaper = False
            .DisablePrinter = False
            .EnablePagePaintHook = False
            .EnablePageSetupHook = False
            .EnablePageSetupTemplate = False
            .EnablePageSetupTemplateHandle = False
            .InHundredthsOfMillimeters = False
            .InThousandthsOfInches = True
            .Margins = True
            .MinMargins = False
            .NoWarning = False
            .ReturnDefault = False
            .ShowHelp = False
        End With

        If vbPageSetupDlg _
                (UserControl.hWnd, _
                m_typ_PageSetupDlg, _
                m_typ_DevMode, _
                m_typ_DevNames, _
                m_typ_PageSetupDlgFlags) Then

            Call pUpdatePageMargins

            ShowPrintDialog = True
        Else
            ShowPrintDialog = False
        End If
    Else
        'Set the Flags
        With m_typ_PrintDlgFlags
            .AllPages = True
            .Collate = False
            .DisablePrintToFile = True
            .EnablePrintHook = False
            .EnablePrintTemplate = False
            .EnablePrintTemplateHandle = False
            .EnableSetupHook = False
            .EnableSetupTemplate = False
            .EnableSetupTemplateHandle = False
            .HidePrintToFile = True
            .NoNetworkButton = False
            .NoPageNums = False
            .NoSelection = True
            .NoWarning = False
            .PageNums = False
            .PrintSetup = False
            .PrintToFile = False
            .ReturnDC = False
            .ReturnDefault = False
            .ReturnIC = False
            .Selection = False
            .ShowHelp = False
            .UseDevModeCopies = False
            .UseDevModeCopiesAndCollate = True
        End With

        m_typ_PrintDlg.nFromPage = 1
        m_typ_PrintDlg.nToPage = Pages.Count
        m_typ_PrintDlg.nMinPage = 1
        m_typ_PrintDlg.nMaxPage = Pages.Count    '&HFFFF

        If vbPrintDlg( _
                UserControl.hWnd, _
                m_typ_PrintDlg, _
                m_typ_DevMode, _
                m_typ_DevNames, _
                m_typ_PrintDlgFlags) Then

            Call pUpdatePageMargins(pdPrinterSetup)

            ShowPrintDialog = True
        Else
            ShowPrintDialog = False
        End If
    End If

    #If RunStackLogger = 1 Then
        Call LogStackItem("ShowPrintDialog - Exit")
    #End If

    Exit Function

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - ShowPrintDialog", Err.Description)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPage,picPage,-1,TextHeight
Public Function TextHeight(ByVal Str As String) As Single
    #If RunStackLogger = 1 Then
        Call LogStackItem("TextHeight")
    #End If

    TextHeight = picPage.TextHeight(Str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picPage,picPage,-1,TextWidth
Public Function TextWidth(ByVal Str As String) As Single
    #If RunStackLogger = 1 Then
        Call LogStackItem("TextWidth")
    #End If

    TextWidth = picPage.TextWidth(Str)
End Function
'*************************************************************

'*************************************************************
' PRIVATE EVENTS
'*************************************************************
Private Sub UserControl_Resize()
    #If RunStackLogger = 1 Then
        Call LogStackItem("UserControl_Resize - Enter")
    #End If

    On Error Resume Next

    #If ShowDebugPrints = 1 Then
        Debug.Print "UserControl_Resize"
    #End If

    Call picWorkspace.Move(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)

    #If RunStackLogger = 1 Then
        Call LogStackItem("UserControl_Resize - Exit")
    #End If
End Sub

Private Sub ScrollBars_Change(eBar As FSScrollBarConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("ScrollBars_Change - Enter")
    #End If

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    #If ShowDebugPrints = 1 Then
        Debug.Print "ScrollBars_Change"
    #End If

    If eBar = fsHorizontal And ScrollBars.Visible(eBar) Then
        Call pMovePage(-Screen.TwipsPerPixelX * ScrollBars.Value(eBar))
    ElseIf ScrollBars.Visible(eBar) Then
        Call pMovePage(picPage.Left, -Screen.TwipsPerPixelY * ScrollBars.Value(eBar))
    End If

    #If RunStackLogger = 1 Then
        Call LogStackItem("ScrollBars_Change - Exit")
    #End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - Scroll_Change", Err.Description)
End Sub

Private Sub ScrollBars_Scroll(eBar As FSScrollBarConstants)
    #If RunStackLogger = 1 Then
        Call LogStackItem("ScrollBars_Scroll - Enter")
    #End If

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    #If ShowDebugPrints = 1 Then
        Debug.Print "ScrollBars_Scroll"
    #End If

    If eBar = fsHorizontal And ScrollBars.Visible(eBar) Then
        Call pMovePage(-Screen.TwipsPerPixelX * ScrollBars.Value(eBar))
    ElseIf ScrollBars.Visible(eBar) Then
        Call pMovePage(picPage.Left, -Screen.TwipsPerPixelY * ScrollBars.Value(eBar))
    End If

    #If RunStackLogger = 1 Then
        Call LogStackItem("ScrollBars_Scroll - Exit")
    #End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - Scroll_Scroll", Err.Description)
End Sub

Private Sub picPage_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    #If RunStackLogger = 1 Then
        Call LogStackItem("picPage_MouseUp")
    #End If

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    Dim sng_HorizPercent As Single
    Dim sng_VertPercent As Single

    If Button = vbLeftButton Then
        'Zoom In
        m_enum_CurrentZoomScale = zs100
        picPage.MouseIcon = LoadResPicture(102, vbResCursor)

        sng_HorizPercent = x / picPage.Width
        sng_VertPercent = Y / picPage.Height

        Call picWorkspace_Resize

        ScrollBars.Value(fsHorizontal) = ScrollBars.Max(fsHorizontal) * sng_HorizPercent
        ScrollBars.Value(fsVertical) = ScrollBars.Max(fsVertical) * sng_VertPercent

    ElseIf Button = vbRightButton And m_enum_CurrentZoomScale <> zsWholePage Then
        'Zoom Out
        m_enum_CurrentZoomScale = zsWholePage
        picPage.MouseIcon = LoadResPicture(101, vbResCursor)

        With ScrollBars
            .Visible(fsVertical) = False
            .Visible(fsHorizontal) = False
        End With

        Call picWorkspace_Resize
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - Page_MouseUp", Err.Description)
End Sub

Private Sub picPage_Paint()
    #If RunStackLogger = 1 Then
        Call LogStackItem("picPage_Paint - Enter")
    #End If

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    #If ShowDebugPrints = 1 Then
        Debug.Print "picPage_Paint"
    #End If

    If Not m_bool_AuthorMode Then
        Call Pages(m_int_CurrentPage).PrintPage
    End If

    #If RunStackLogger = 1 Then
        Call LogStackItem("picPage_Paint - Exit")
    #End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - picPage_Paint", Err.Description)
End Sub

Private Sub picWorkspace_Resize()
    #If RunStackLogger = 1 Then
        Call LogStackItem("picWorkspace_Resize - Enter")
    #End If

    On Error Resume Next

    Dim lng_Height  As Long
    Dim lng_Width   As Long

    #If ShowDebugPrints = 1 Then
        Debug.Print "picWorkspace_Resize"
    #End If

    If Printers.Count > 0 Then
        Select Case m_enum_CurrentZoomScale
            Case zsWholePage
                lng_Width = ((picWorkspace.ScaleHeight - m_int_PaperShadowOffset - m_cint_WkspPgMargin * 2) * m_lng_PaperWidth) / m_lng_PaperHeight
                If lng_Width <= (picWorkspace.ScaleWidth - m_cint_WkspPgMargin * 2) Then
                    Call pMovePage( _
                            (picWorkspace.ScaleWidth - lng_Width) / 2, _
                            m_cint_WkspPgMargin, _
                            lng_Width, _
                            (picWorkspace.ScaleHeight - m_int_PaperShadowOffset - m_cint_WkspPgMargin * 2))
                Else
                    lng_Height = ((picWorkspace.ScaleWidth - m_int_PaperShadowOffset - m_cint_WkspPgMargin * 2) * m_lng_PaperHeight) / m_lng_PaperWidth
                    Call pMovePage( _
                            m_cint_WkspPgMargin, _
                            ((picWorkspace.ScaleHeight - m_int_PaperShadowOffset - lng_Height) / 2), _
                            (picWorkspace.ScaleWidth - m_cint_WkspPgMargin * 2), _
                            lng_Height)
                End If
            Case zs100
                If picPage.Width <= picWorkspace.ScaleWidth And picPage.Height <= picWorkspace.ScaleHeight Then
                    Call pMovePage((picWorkspace.ScaleWidth - m_lng_PaperWidth) / 2, (picWorkspace.ScaleHeight - m_lng_PaperHeight) / 2, m_lng_PaperWidth, m_lng_PaperHeight)
                ElseIf picPage.Width > picWorkspace.ScaleWidth And picPage.Height <= picWorkspace.ScaleHeight Then
                    Call pMovePage(0, (picWorkspace.ScaleHeight - m_lng_PaperHeight) / 2, m_lng_PaperWidth, m_lng_PaperHeight)
                ElseIf picPage.Width <= picWorkspace.ScaleWidth And picPage.Height > picWorkspace.ScaleHeight Then
                    Call pMovePage((picWorkspace.ScaleWidth - m_lng_PaperWidth) / 2, 0, m_lng_PaperWidth, m_lng_PaperHeight)
                ElseIf picPage.Width > picWorkspace.ScaleWidth And picPage.Height > picWorkspace.ScaleHeight Then
                    Call pMovePage(0, 0, m_lng_PaperWidth, m_lng_PaperHeight)
                End If

                With ScrollBars
                    If picPage.Height > picWorkspace.ScaleHeight Then
                        .Max(fsVertical) = (picPage.Height - picWorkspace.ScaleHeight) \ Screen.TwipsPerPixelY
                        .LargeChange(fsVertical) = (picPage.Height / 10) \ Screen.TwipsPerPixelY
                        .SmallChange(fsVertical) = (picPage.Height / 20) \ Screen.TwipsPerPixelY
                        .Visible(fsVertical) = True
                    Else
                        .Visible(fsVertical) = False
                    End If

                    If picPage.Width > picWorkspace.ScaleWidth Then
                        .Max(fsHorizontal) = (picPage.Width - picWorkspace.ScaleWidth) \ Screen.TwipsPerPixelX
                        .LargeChange(fsHorizontal) = (picPage.Width / 10) \ Screen.TwipsPerPixelX
                        .SmallChange(fsHorizontal) = (picPage.Width / 20) \ Screen.TwipsPerPixelX
                        .Visible(fsHorizontal) = True
                    Else
                        .Visible(fsHorizontal) = False
                    End If
                End With
        End Select
    Else
        Call imgNoPrinters.Move( _
                (picWorkspace.ScaleWidth / 2) - (imgNoPrinters.Width / 2), _
                (picWorkspace.ScaleHeight / 2) - (imgNoPrinters.Height))

        Call lblNoPrinters.Move( _
                (picWorkspace.ScaleWidth / 2) - (lblNoPrinters.Width / 2), _
                picWorkspace.ScaleHeight / 2)
    End If

    #If RunStackLogger = 1 Then
        Call LogStackItem("picWorkspace_Resize - Exit")
    #End If
End Sub

Private Sub UserControl_InitProperties()
    #If RunStackLogger = 1 Then
        Call LogStackItem("UserControl_InitProperties")
    #End If

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    #If ShowDebugPrints = 1 Then
        Debug.Print "UserControl_InitProperties"
    #End If

    AbortWindow = True
    UserControl.Appearance = ap3D
    picWorkspace.BackColor = vbApplicationWorkspace
    UserControl.BorderStyle = bsSingle
    DefaultDevice = True
    DrawStyle = dsSolid
    DrawWidth = 1
    UserControl.Enabled = True
    FileName = ""
    picPage.FillColor = vbWhite
    picPage.FillStyle = fsTransparent
    picPage.ForeColor = vbBlack
    picPage.Font = picPage.Font
    picPage.FontBold = picPage.FontBold
    picPage.FontItalic = picPage.FontItalic
    picPage.FontName = picPage.FontName
    picPage.FontSize = picPage.FontSize
    picPage.FontStrikethru = picPage.FontStrikethru
    picPage.FontTransparent = picPage.FontTransparent
    picPage.FontUnderline = picPage.FontUnderline
    MarginBottom = 1440
    MarginLeft = 1440
    MarginRight = 1440
    MarginTop = 1440
    Set MouseIcon = Nothing
    MouseZoom = mzSimple
    PageBorder = pbNone
    PaperBorderColor = vbBlack
    PaperBorderWidth = 1
    PaperShadowBorderStyle = bsSingle
    PaperShadowColor = vbBlack
    PaperShadowOffset = 20
    PhysicalPage = False
    Preview = True
    ShowGuides = sgHide
    Zoom = 0
    ZoomMax = 400
    ZoomMin = 10
    ZoomStep = 10

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - UserControl_InitProperties", Err.Description)
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    #If RunStackLogger = 1 Then
        Call LogStackItem("UserControl_ReadProperties")
    #End If

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    #If ShowDebugPrints = 1 Then
        Debug.Print "UserControl_ReadProperties"
    #End If

    AbortWindow = PropBag.ReadProperty("AbortWindow", True)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", ap3D)
    picWorkspace.BackColor = PropBag.ReadProperty("BackColor", vbApplicationWorkspace)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", bsSingle)
    DefaultDevice = PropBag.ReadProperty("DefaultDevice", True)
    DrawStyle = PropBag.ReadProperty("DrawStyle", dsSolid)
    DrawWidth = PropBag.ReadProperty("DrawWidth", 1)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    FileName = PropBag.ReadProperty("FileName", "")
    picPage.FillColor = PropBag.ReadProperty("FillColor", vbWhite)
    picPage.FillStyle = PropBag.ReadProperty("FillStyle", fsTransparent)
    picPage.ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
    picPage.Font = PropBag.ReadProperty("Font", picPage.Font)
    picPage.FontBold = PropBag.ReadProperty("FontBold", picPage.FontBold)
    picPage.FontItalic = PropBag.ReadProperty("FontItalic", picPage.FontItalic)
    picPage.FontName = PropBag.ReadProperty("FontName", picPage.FontName)
    picPage.FontSize = PropBag.ReadProperty("FontSize", picPage.FontSize)
    picPage.FontStrikethru = PropBag.ReadProperty("FontStrikethru", picPage.FontStrikethru)
    picPage.FontTransparent = PropBag.ReadProperty("FontTransparent", picPage.FontTransparent)
    picPage.FontUnderline = PropBag.ReadProperty("FontUnderline", picPage.FontUnderline)
    MarginBottom = PropBag.ReadProperty("MarginBottom", 1440)
    MarginLeft = PropBag.ReadProperty("MarginLeft", 1440)
    MarginRight = PropBag.ReadProperty("MarginRight", 1440)
    MarginTop = PropBag.ReadProperty("MarginTop", 1440)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", mpDefault)
    MouseZoom = PropBag.ReadProperty("MouseZoom", mzSimple)
    PageBorder = PropBag.ReadProperty("PageBorder", pbNone)
    PaperBorderColor = PropBag.ReadProperty("PaperBorderColor", vbBlack)
    PaperBorderWidth = PropBag.ReadProperty("PaperBorderWidth", 1)
    PaperShadowBorderStyle = PropBag.ReadProperty("PaperShadowBorderStyle", bsSingle)
    PaperShadowColor = PropBag.ReadProperty("PaperShadowColor", vbBlack)
    PaperShadowOffset = PropBag.ReadProperty("PaperShadowOffset", 20)
    PhysicalPage = PropBag.ReadProperty("PhysicalPage", False)
    Preview = PropBag.ReadProperty("Preview", True)
    ShowGuides = PropBag.ReadProperty("ShowGuides", sgHide)
    Zoom = PropBag.ReadProperty("Zoom", 0)
    ZoomMax = PropBag.ReadProperty("Zoom", 400)
    ZoomMin = PropBag.ReadProperty("Zoom", 10)
    ZoomStep = PropBag.ReadProperty("ZoomStep", 10)

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - UserControl_ReadProperties", Err.Description)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    #If RunStackLogger = 1 Then
        Call LogStackItem("UserControl_WriteProperties")
    #End If

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    #If ShowDebugPrints = 1 Then
        Debug.Print "UserControl_WriteProperties"
    #End If

    Call PropBag.WriteProperty("AbortWindow", AbortWindow, True)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, ap3D)
    Call PropBag.WriteProperty("BackColor", picWorkspace.BackColor, vbApplicationWorkspace)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, bsSingle)
    Call PropBag.WriteProperty("DefaultDevice", DefaultDevice, True)
    Call PropBag.WriteProperty("DrawStyle", picPage.DrawStyle, 0)
    Call PropBag.WriteProperty("DrawWidth", picPage.DrawWidth, 1)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("FileName", FileName, "")
    Call PropBag.WriteProperty("FillColor", picPage.FillColor, vbWhite)
    Call PropBag.WriteProperty("FillStyle", picPage.FillStyle, fsTransparent)
    Call PropBag.WriteProperty("ForeColor", picPage.ForeColor, vbBlack)
    Call PropBag.WriteProperty("Font", picPage.Font, picPage.Font)
    Call PropBag.WriteProperty("FontBold", picPage.FontBold, picPage.FontBold)
    Call PropBag.WriteProperty("FontItalic", picPage.FontItalic, picPage.FontItalic)
    Call PropBag.WriteProperty("FontName", picPage.FontName, picPage.FontName)
    Call PropBag.WriteProperty("FontSize", picPage.FontSize, picPage.FontSize)
    Call PropBag.WriteProperty("FontStrikethru", picPage.FontStrikethru, picPage.FontStrikethru)
    Call PropBag.WriteProperty("FontTransparent", picPage.FontTransparent, True)
    Call PropBag.WriteProperty("FontUnderline", picPage.FontUnderline, picPage.FontUnderline)
    Call PropBag.WriteProperty("MarginBottom", MarginBottom, 1440)
    Call PropBag.WriteProperty("MarginLeft", MarginLeft, 1440)
    Call PropBag.WriteProperty("MarginRight", MarginRight, 1440)
    Call PropBag.WriteProperty("MarginTop", MarginTop, 1440)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, UserControl.MousePointer)
    Call PropBag.WriteProperty("MouseZoom", MouseIcon, mzSimple)
    Call PropBag.WriteProperty("PageBorder", PageBorder, pbNone)
    Call PropBag.WriteProperty("PaperBorderColor", PaperBorderColor, vbBlack)
    Call PropBag.WriteProperty("PaperBorderWidth", PaperBorderWidth, 1)
    Call PropBag.WriteProperty("PaperShadowBorderStyle", PaperShadowBorderStyle, bsSingle)
    Call PropBag.WriteProperty("PaperShadowColor", PaperShadowColor, vbBlack)
    Call PropBag.WriteProperty("PaperShadowOffset", PaperShadowOffset, 20)
    Call PropBag.WriteProperty("PhysicalPage", PhysicalPage, False)
    Call PropBag.WriteProperty("Preview", Preview, True)
    Call PropBag.WriteProperty("ShowGuides", ShowGuides, sgHide)
    Call PropBag.WriteProperty("Zoom", Zoom, 0)
    Call PropBag.WriteProperty("ZoomMax", ZoomMax, 400)
    Call PropBag.WriteProperty("ZoomMin", ZoomMin, 10)
    Call PropBag.WriteProperty("ZoomStep", ZoomStep, 10)

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - UserControl_WriteProperties", Err.Description)
End Sub
'*************************************************************

'*************************************************************
' PRIVATE FUNCTIONS/SUBROUTINES
'*************************************************************
Private Sub pInitPageParams()
    #If RunStackLogger = 1 Then
        Call LogStackItem("pInitPageParams")
    #End If

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    #If ShowDebugPrints = 1 Then
        Debug.Print "m_lng_DeviceResolutionX: " & m_lng_DeviceResolutionX,
    #End If
    m_lng_DeviceResolutionX = GetDeviceCaps(Printer.hdc, LOGPIXELSX)
    #If ShowDebugPrints = 1 Then
        Debug.Print m_lng_DeviceResolutionX
        Debug.Print "m_lng_DeviceResolutionY: " & m_lng_DeviceResolutionY,
    #End If
    m_lng_DeviceResolutionY = GetDeviceCaps(Printer.hdc, LOGPIXELSY)
    #If ShowDebugPrints = 1 Then
        Debug.Print m_lng_DeviceResolutionY
        Debug.Print "m_lng_MarginLeft: " & m_lng_MarginLeft,
    #End If
    m_lng_MarginLeft = (GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX) / m_lng_DeviceResolutionX) * m_cint_TwipsPerInch
    #If ShowDebugPrints = 1 Then
        Debug.Print m_lng_MarginLeft
        Debug.Print "m_lng_MarginTop: " & m_lng_MarginTop,
    #End If
    m_lng_MarginTop = (GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY) / m_lng_DeviceResolutionY) * m_cint_TwipsPerInch
    #If ShowDebugPrints = 1 Then
        Debug.Print m_lng_MarginTop
        Debug.Print "m_lng_PageWidth: " & m_lng_PageWidth,
    #End If
    m_lng_PageWidth = (GetDeviceCaps(Printer.hdc, HORZRES) / m_lng_DeviceResolutionX) * m_cint_TwipsPerInch
    #If ShowDebugPrints = 1 Then
        Debug.Print m_lng_PageWidth
        Debug.Print "m_lng_PageHeight: " & m_lng_PageHeight,
    #End If
    m_lng_PageHeight = (GetDeviceCaps(Printer.hdc, VERTRES) / m_lng_DeviceResolutionY) * m_cint_TwipsPerInch
    #If ShowDebugPrints = 1 Then
        Debug.Print m_lng_PageHeight
        Debug.Print "m_lng_PaperWidth: " & m_lng_PaperWidth,
    #End If
    m_lng_PaperWidth = (GetDeviceCaps(Printer.hdc, PHYSICALWIDTH) / m_lng_DeviceResolutionX) * m_cint_TwipsPerInch
    #If ShowDebugPrints = 1 Then
        Debug.Print m_lng_PaperWidth
        Debug.Print "m_lng_MarginRight: " & m_lng_MarginRight,
    #End If
    m_lng_MarginRight = ((m_lng_PaperWidth - m_lng_PageWidth - m_lng_MarginLeft) / m_lng_DeviceResolutionX) * m_cint_TwipsPerInch
    #If ShowDebugPrints = 1 Then
        Debug.Print m_lng_MarginRight
        Debug.Print "m_lng_PaperHeight: " & m_lng_PaperHeight,
    #End If
    m_lng_PaperHeight = (GetDeviceCaps(Printer.hdc, PHYSICALHEIGHT) / m_lng_DeviceResolutionY) * m_cint_TwipsPerInch
    #If ShowDebugPrints = 1 Then
        Debug.Print m_lng_PaperHeight
        Debug.Print "m_lng_MarginBottom: " & m_lng_MarginBottom,
    #End If
    m_lng_MarginBottom = ((m_lng_PaperHeight - m_lng_PageHeight - m_lng_MarginTop) / m_lng_DeviceResolutionY) * m_cint_TwipsPerInch
    #If ShowDebugPrints = 1 Then
        Debug.Print m_lng_MarginBottom
    #End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - pInitPageParams", Err.Description)
End Sub

Private Sub pMovePage(Left As Single, Optional Top, Optional Width, Optional Height)
    #If RunStackLogger = 1 Then
        Call LogStackItem("pMovePage - Enter")
    #End If

    On Error Resume Next

    If IsMissing(Top) And IsMissing(Width) And IsMissing(Height) Then    '000
        Call picPage.Move(Left)
        Call picPageShadow.Move(Left + m_int_PaperShadowOffset)

    ElseIf IsMissing(Top) And IsMissing(Width) And Not IsMissing(Height) Then    '001
        If m_int_Height <> Height Then
            Call picPage.Move(Left, , , CSng(Height))
            Call shpPaperBorder.Move(0, 0, picPage.ScaleWidth, picPage.ScaleHeight)
            Call picPageShadow.Move(Left + m_int_PaperShadowOffset, , , CSng(Height))
            m_int_Height = Height
        ElseIf picPage.Left <> Left Then
            Call picPage.Move(Left)
        End If

    ElseIf IsMissing(Top) And Not IsMissing(Width) And IsMissing(Height) Then    '010
        If m_int_Width <> Width Then
            Call picPage.Move(Left, , CSng(Width))
            Call shpPaperBorder.Move(0, 0, picPage.ScaleWidth, picPage.ScaleHeight)
            Call picPageShadow.Move(Left + m_int_PaperShadowOffset, , CSng(Width))
            m_int_Width = Width
        ElseIf picPage.Left <> Left Then
            Call picPage.Move(Left)
        End If

    ElseIf Not IsMissing(Top) And Not IsMissing(Width) And Not IsMissing(Height) Then    '111
        If m_int_Width <> Width Or m_int_Height <> Height Then
            Call picPage.Move(Left, CSng(Top), CSng(Width), CSng(Height))
            Call shpPaperBorder.Move(0, 0, picPage.ScaleWidth, picPage.ScaleHeight)
            Call picPageShadow.Move(Left + m_int_PaperShadowOffset, CSng(Top) + m_int_PaperShadowOffset, CSng(Width), CSng(Height))
            m_int_Height = Height
            m_int_Width = Width
        ElseIf picPage.Left <> Left Or picPage.Left <> Top Then
            Call picPage.Move(Left, CSng(Top))
            Call picPageShadow.Move(Left + m_int_PaperShadowOffset, CSng(Top) + m_int_PaperShadowOffset)
        End If

    ElseIf Not IsMissing(Top) And IsMissing(Width) And IsMissing(Height) Then    '100
        Call picPage.Move(Left, CSng(Top))
        Call picPageShadow.Move(Left + m_int_PaperShadowOffset, CSng(Top) + m_int_PaperShadowOffset)

    ElseIf Not IsMissing(Top) And IsMissing(Width) And Not IsMissing(Height) Then    '101
        If m_int_Height <> Height Then
            Call picPage.Move(Left, CSng(Top), , CSng(Height))
            Call shpPaperBorder.Move(0, 0, picPage.ScaleWidth, picPage.ScaleHeight)
            Call picPageShadow.Move(Left + m_int_PaperShadowOffset, CSng(Top) + m_int_PaperShadowOffset, , CSng(Height))
            m_int_Height = Height
        ElseIf picPage.Left <> Left Or picPage.Left <> Top Then
            Call picPage.Move(Left, CSng(Top))
            Call picPageShadow.Move(Left + m_int_PaperShadowOffset, CSng(Top) + m_int_PaperShadowOffset)
        End If

    End If

    #If RunStackLogger = 1 Then
        Call LogStackItem("pMovePage - Exit")
    #End If
End Sub

Private Sub pValidateFromToPages(FromPage As Long, ToPage As Long)
    #If RunStackLogger = 1 Then
        Call LogStackItem("pValidateFromToPages")
    #End If

    If FromPage <= 0 Or FromPage > Pages.Count Then FromPage = 1
    If ToPage < FromPage Or ToPage > Pages.Count Then ToPage = Pages.Count
End Sub

Private Sub pUpdatePageMargins(Optional DialogType As PrintDialogConstants = pdPageSetup)
    #If RunStackLogger = 1 Then
        Call LogStackItem("pUpdatePageMargins")
    #End If

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    Dim obj_Pages   As Page

    If DialogType = pdPageSetup Then
        m_lng_PaperWidth = pConvertPrinterUnitsToTwips(m_typ_PageSetupDlg.ptPaperSize.x)
        m_lng_PaperHeight = pConvertPrinterUnitsToTwips(m_typ_PageSetupDlg.ptPaperSize.Y)
    Else
        If m_typ_DevMode.dmOrientation = 1 Then
            m_lng_PaperWidth = ScaleX((m_typ_DevMode.dmPaperWidth / 10), vbMillimeters, vbTwips)
            m_lng_PaperHeight = ScaleY((m_typ_DevMode.dmPaperLength / 10), vbMillimeters, vbTwips)
        Else
            m_lng_PaperHeight = ScaleX((m_typ_DevMode.dmPaperWidth / 10), vbMillimeters, vbTwips)
            m_lng_PaperWidth = ScaleY((m_typ_DevMode.dmPaperLength / 10), vbMillimeters, vbTwips)
        End If
    End If
    m_lng_MarginBottom = pConvertPrinterUnitsToTwips(m_typ_PageSetupDlg.rtMargin.bottom)
    m_lng_MarginLeft = pConvertPrinterUnitsToTwips(m_typ_PageSetupDlg.rtMargin.Left)
    m_lng_MarginRight = pConvertPrinterUnitsToTwips(m_typ_PageSetupDlg.rtMargin.right)
    m_lng_MarginTop = pConvertPrinterUnitsToTwips(m_typ_PageSetupDlg.rtMargin.Top)
    m_lng_PageHeight = m_lng_PaperHeight - m_lng_MarginTop - m_lng_MarginBottom
    m_lng_PageWidth = m_lng_PaperWidth - m_lng_MarginLeft - m_lng_MarginRight

    #If ShowDebugPrints = 1 Then
        Debug.Print "***************Setting***************"
        Debug.Print "PaperWidth: ", m_lng_PaperWidth
        Debug.Print "PaperHeight: ", m_lng_PaperHeight
        Debug.Print "MarginBottom: ", m_lng_MarginBottom
        Debug.Print "MarginLeft: ", m_lng_MarginLeft
        Debug.Print "MarginRight: ", m_lng_MarginRight
        Debug.Print "MarginTop: ", m_lng_MarginTop
        Debug.Print "PageHeight: ", m_lng_PageHeight
        Debug.Print "PageWidth: ", m_lng_PageWidth
        Debug.Print "***************Done***************"
    #End If

    For Each obj_Pages In Pages
        obj_Pages.MarginBottom = m_lng_MarginBottom
        obj_Pages.MarginLeft = m_lng_MarginLeft
        obj_Pages.MarginRight = m_lng_MarginRight
        obj_Pages.MarginTop = m_lng_MarginTop
        obj_Pages.PaperHeight = m_lng_PaperHeight
        obj_Pages.PaperWidth = m_lng_PaperWidth
        obj_Pages.PageHeight = m_lng_PageHeight
        obj_Pages.PageWidth = m_lng_PageWidth
    Next

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - pInitPageParams", Err.Description)
End Sub

Private Function pConvertTwipsToInches(in_Twips As Long) As Single
    #If RunStackLogger = 1 Then
        Call LogStackItem("pConvertTwipsToInches")
    #End If

    pConvertTwipsToInches = in_Twips / m_cint_TwipsPerInch
End Function

Private Function pConvertInchesToTwips(in_Inches As Single) As Long
    #If RunStackLogger = 1 Then
        Call LogStackItem("pConvertInchesToTwips")
    #End If

    pConvertInchesToTwips = in_Inches * m_cint_TwipsPerInch
End Function

Private Function pConvertTwipsToPrinterUnits(in_Twips As Long) As Long
    #If RunStackLogger = 1 Then
        Call LogStackItem("pConvertTwipsToPrinterUnits")
    #End If

    pConvertTwipsToPrinterUnits = (in_Twips / m_cint_TwipsPerInch) * m_clng_Units
End Function

Private Function pConvertInchesToPrinterUnits(in_Inches As Single) As Long
    #If RunStackLogger = 1 Then
        Call LogStackItem("pConvertInchesToPrinterUnits")
    #End If

    pConvertInchesToPrinterUnits = in_Inches * m_clng_Units
End Function

Private Function pConvertPrinterUnitsToTwips(in_Units As Long) As Long
    #If RunStackLogger = 1 Then
        Call LogStackItem("pConvertPrinterUnitsToTwips")
    #End If

    pConvertPrinterUnitsToTwips = (in_Units / m_clng_Units) * m_cint_TwipsPerInch
End Function














