Attribute VB_Name = "modCommonDialog"
Option Explicit

Global g_int_CurrentPrinterIndex    As Integer

Public Enum PrinterUnitsConstants
    puHundredthsOfMillimeters = PSD_INHUNDREDTHSOFMILLIMETERS
    puThousandthsOfInches = PSD_INTHOUSANDTHSOFINCHES
End Enum

Public Type PAGESETUPDLGFLAGS
    DefaultMinMargins               As Boolean
    DisableMargins                  As Boolean
    DisableOrientation              As Boolean
    DisablePagePainting             As Boolean
    DisablePaper                    As Boolean
    DisablePrinter                  As Boolean
    EnablePagePaintHook             As Boolean
    EnablePageSetupHook             As Boolean
    EnablePageSetupTemplate         As Boolean
    EnablePageSetupTemplateHandle   As Boolean
    InHundredthsOfMillimeters       As Boolean
    InThousandthsOfInches           As Boolean
    Margins                         As Boolean
    MinMargins                      As Boolean
    NoWarning                       As Boolean
    ReturnDefault                   As Boolean
    ShowHelp                        As Boolean
End Type

Public Type PRINTDLGFLAGS
    AllPages                        As Boolean
    Collate                         As Boolean
    DisablePrintToFile              As Boolean
    EnablePrintHook                 As Boolean
    EnablePrintTemplate             As Boolean
    EnablePrintTemplateHandle       As Boolean
    EnableSetupHook                 As Boolean
    EnableSetupTemplate             As Boolean
    EnableSetupTemplateHandle       As Boolean
    HidePrintToFile                 As Boolean
    NoNetworkButton                 As Boolean
    NoPageNums                      As Boolean
    NoSelection                     As Boolean
    NoWarning                       As Boolean
    PageNums                        As Boolean
    PrintSetup                      As Boolean
    PrintToFile                     As Boolean
    ReturnDC                        As Boolean
    ReturnDefault                   As Boolean
    ReturnIC                        As Boolean
    Selection                       As Boolean
    ShowHelp                        As Boolean
    UseDevModeCopies                As Boolean
    UseDevModeCopiesAndCollate      As Boolean
End Type

Public Type DEVMODEFIELDS
    InitOrientation                 As Boolean
    InitPaperSize                   As Boolean
    InitPaperLenght                 As Boolean
    InitPaperWidth                  As Boolean
    InitPosition                    As Boolean    'Windows 98, Windows NT 5.0 and later
    InitScale                       As Boolean
    InitCopies                      As Boolean
    InitDefaultSource               As Boolean
    InitPrintQuality                As Boolean
    InitColor                       As Boolean
    InitDuplex                      As Boolean
    InitYResolution                 As Boolean
    InitTToption                    As Boolean
    InitCollate                     As Boolean
    InitFormname                    As Boolean
    InitLogPixels                   As Boolean
    InitBitsPerPel                  As Boolean
    InitPelsWidth                   As Boolean
    InitPelsHeight                  As Boolean
    InitDisplayFlags                As Boolean
    InitDisplayFrequency            As Boolean
    InitICMMethod                   As Boolean
    InitICMIntent                   As Boolean
    InitMediaType                   As Boolean
    InitDitherType                  As Boolean
    InitPanningWidth                As Boolean    'Windows NT 5.0 and later: dmPanningWidth
    InitPanningHeight               As Boolean    'Windows NT 5.0 and later:
End Type

'By placing these variables here we can retain their data when the user re-visits the dlg
Private ISetDfltPrinter         As New cSetDfltPrinter
Private m_lng_DevMode           As Long
Private m_lng_DevNames          As Long

' PrintDialog wrapper
Public Function vbPrintDlg( _
    ByRef hWndOwner As Long, _
    ByRef typ_PrintDlg As TPRINTDLG, _
    ByRef typ_DevMode As DEVMODE, _
    ByRef typ_DevNames As DEVNAMES, _
    ByRef typ_PrintDlgFlags As PRINTDLGFLAGS _
    ) As Boolean

    Dim obj_SelPrinter          As Printer
    Dim str_NewPrinterName      As String
    Dim hResult As Long

    ' Fill in TPRINTDLG structure
    typ_PrintDlg.lStructSize = Len(typ_PrintDlg)
    typ_PrintDlg.hWndOwner = hWndOwner

    'Set the Flags
    typ_PrintDlg.flags = _
            -typ_PrintDlgFlags.Collate * PD_COLLATE Or _
            -typ_PrintDlgFlags.DisablePrintToFile * PD_DISABLEPRINTTOFILE Or _
            -typ_PrintDlgFlags.EnablePrintHook * PD_ENABLEPRINTHOOK Or _
            -typ_PrintDlgFlags.EnablePrintTemplate * PD_ENABLEPRINTTEMPLATE Or _
            -typ_PrintDlgFlags.EnablePrintTemplateHandle * PD_ENABLEPRINTTEMPLATEHANDLE Or _
            -typ_PrintDlgFlags.EnableSetupHook * PD_ENABLESETUPHOOK Or _
            -typ_PrintDlgFlags.EnableSetupTemplate * PD_ENABLESETUPTEMPLATE Or _
            -typ_PrintDlgFlags.EnableSetupTemplateHandle * PD_ENABLESETUPTEMPLATEHANDLE Or _
            -typ_PrintDlgFlags.HidePrintToFile * PD_HIDEPRINTTOFILE Or _
            -typ_PrintDlgFlags.NoNetworkButton * PD_NONETWORKBUTTON Or _
            -typ_PrintDlgFlags.NoPageNums * PD_NOPAGENUMS Or _
            -typ_PrintDlgFlags.NoSelection * PD_NOSELECTION Or _
            -typ_PrintDlgFlags.NoWarning * PD_NOWARNING Or _
            -typ_PrintDlgFlags.PrintSetup * PD_PRINTSETUP Or _
            -typ_PrintDlgFlags.PrintToFile * PD_PRINTTOFILE Or _
            -typ_PrintDlgFlags.ReturnDC * PD_RETURNDC Or _
            -typ_PrintDlgFlags.ReturnDefault * PD_RETURNDEFAULT Or _
            -typ_PrintDlgFlags.ReturnIC * PD_RETURNIC Or _
            -typ_PrintDlgFlags.ShowHelp * PD_SHOWHELP Or _
            -typ_PrintDlgFlags.UseDevModeCopies * PD_USEDEVMODECOPIES Or _
            -typ_PrintDlgFlags.UseDevModeCopiesAndCollate * PD_USEDEVMODECOPIESANDCOLLATE

    If typ_PrintDlgFlags.AllPages Then
        typ_PrintDlg.flags = typ_PrintDlg.flags Or PD_ALLPAGES

    ElseIf typ_PrintDlgFlags.PageNums Then
        typ_PrintDlg.flags = typ_PrintDlg.flags Or PD_PAGENUMS

    ElseIf typ_PrintDlgFlags.Selection Then
        typ_PrintDlg.flags = typ_PrintDlg.flags Or PD_SELECTION

    End If

    'Set DEVMODE structure from TPRINTDLG
    m_lng_DevMode = GlobalLock(typ_PrintDlg.hDevMode)
    If m_lng_DevMode > 0 Then
        Call CopyMemory(ByVal m_lng_DevMode, typ_DevMode, Len(typ_DevMode))
        hResult = GlobalUnlock(typ_PrintDlg.hDevMode)
    End If

    'Set DEVNAMES structure from TPRINTDLG
    m_lng_DevNames = GlobalLock(typ_PrintDlg.hDevNames)
    If m_lng_DevNames > 0 Then
        Call CopyMemory(ByVal m_lng_DevNames, typ_DevNames, Len(typ_DevNames))
        hResult = GlobalUnlock(typ_PrintDlg.hDevNames)
    End If

    ' Show Print dialog
    If PrintDlg(typ_PrintDlg) Then
        vbPrintDlg = True

        'Call Hourglass(hWndOwner, True)

        If Not GetPrinterStructs(typ_PrintDlg.hDevMode, typ_PrintDlg.hDevNames, typ_DevMode, typ_DevNames, True) Then
            vbPrintDlg = False
        End If

        'Call Hourglass(hWndOwner, False)
    Else
        vbPrintDlg = False
    End If
End Function

' PageSetupDlg wrapper
Public Function vbPageSetupDlg( _
    ByRef hWndOwner As Long, _
    ByRef typ_PrintSetupDlg As TPAGESETUPDLG, _
    ByRef typ_DevMode As DEVMODE, _
    ByRef typ_DevNames As DEVNAMES, _
    ByRef typ_PageSetupDlgFlags As PAGESETUPDLGFLAGS _
    ) As Boolean

    Dim hResult As Long

    ' Fill in TPRINTDLG structure
    typ_PrintSetupDlg.lStructSize = Len(typ_PrintSetupDlg)
    typ_PrintSetupDlg.hWndOwner = hWndOwner

    'Set the Flags
    typ_PrintSetupDlg.flags = _
            -typ_PageSetupDlgFlags.DefaultMinMargins * PSD_DEFAULTMINMARGINS Or _
            -typ_PageSetupDlgFlags.DisableMargins * PSD_DISABLEMARGINS Or _
            -typ_PageSetupDlgFlags.DisableOrientation * PSD_DISABLEORIENTATION Or _
            -typ_PageSetupDlgFlags.DisablePagePainting * PSD_DISABLEPAGEPAINTING Or _
            -typ_PageSetupDlgFlags.DisablePaper * PSD_DISABLEPAPER Or _
            -typ_PageSetupDlgFlags.DisablePrinter * PSD_DISABLEPRINTER Or _
            -typ_PageSetupDlgFlags.EnablePagePaintHook * PSD_ENABLEPAGEPAINTHOOK Or _
            -typ_PageSetupDlgFlags.EnablePageSetupHook * PSD_ENABLEPAGESETUPHOOK Or _
            -typ_PageSetupDlgFlags.EnablePageSetupTemplate * PSD_ENABLEPAGESETUPTEMPLATE Or _
            -typ_PageSetupDlgFlags.EnablePageSetupTemplateHandle * PSD_ENABLEPAGESETUPTEMPLATEHANDLE Or _
            -typ_PageSetupDlgFlags.Margins * PSD_MARGINS Or _
            -typ_PageSetupDlgFlags.MinMargins * PSD_MINMARGINS Or _
            -typ_PageSetupDlgFlags.NoWarning * PSD_NOWARNING Or _
            -typ_PageSetupDlgFlags.ReturnDefault * PSD_RETURNDEFAULT Or _
            -typ_PageSetupDlgFlags.ShowHelp * PSD_SHOWHELP

    If typ_PageSetupDlgFlags.InHundredthsOfMillimeters Then
        typ_PrintSetupDlg.flags = typ_PrintSetupDlg.flags Or PSD_INHUNDREDTHSOFMILLIMETERS
    Else
        typ_PrintSetupDlg.flags = typ_PrintSetupDlg.flags Or PSD_INTHOUSANDTHSOFINCHES
    End If

    'Set DEVMODE structure from TPRINTDLG
    m_lng_DevMode = GlobalLock(typ_PrintSetupDlg.hDevMode)
    If m_lng_DevMode > 0 Then
        Call CopyMemory(ByVal m_lng_DevMode, typ_DevMode, Len(typ_DevMode))
        hResult = GlobalUnlock(typ_PrintSetupDlg.hDevMode)
    End If

    'Set DEVNAMES structure from TPRINTDLG
    m_lng_DevNames = GlobalLock(typ_PrintSetupDlg.hDevNames)
    If m_lng_DevNames > 0 Then
        Call CopyMemory(ByVal m_lng_DevNames, typ_DevNames, Len(typ_DevNames))
        hResult = GlobalUnlock(typ_PrintSetupDlg.hDevNames)
    End If

    ' Show Print dialog
    If PageSetupDlg(typ_PrintSetupDlg) Then
        vbPageSetupDlg = True

        If Not GetPrinterStructs(typ_PrintSetupDlg.hDevMode, typ_PrintSetupDlg.hDevNames, typ_DevMode, typ_DevNames, True) Then
            vbPageSetupDlg = False
        End If

    End If
End Function

'PageSetupDlg wrapper
Public Function vbInitPageSetupDlg( _
    ByRef hWndOwner As Long, _
    ByRef typ_PrintSetupDlg As TPAGESETUPDLG, _
    ByRef typ_PrintDlg As TPRINTDLG, _
    ByRef typ_DevMode As DEVMODE, _
    ByRef typ_DevNames As DEVNAMES, _
    ByRef enum_PrinterUnits As PrinterUnitsConstants _
    ) As Boolean

    Dim hResult As Long

    'Set the Flags
    typ_PrintSetupDlg.lStructSize = Len(typ_PrintSetupDlg)
    If enum_PrinterUnits = puThousandthsOfInches Then
        typ_PrintSetupDlg.flags = PSD_RETURNDEFAULT Or PSD_INTHOUSANDTHSOFINCHES
    Else
        typ_PrintSetupDlg.flags = PSD_RETURNDEFAULT Or PSD_INHUNDREDTHSOFMILLIMETERS
    End If
    
    '**************************************
    ' The following code produces a Dr. Watson visit
    ' in certain versions of Windows... I still
    ' never understood why.
    '**************************************

    '    'Memory allocation must be done...
    '    typ_PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(typ_DevMode))
    '    typ_PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(typ_DevNames))
    '
    '    'Set DEVMODE structure from TPRINTDLG
    '    m_lng_DevMode = GlobalLock(typ_PrintDlg.hDevMode)
    '    If m_lng_DevMode > 0 Then
    '        Call CopyMemory(ByVal m_lng_DevMode, typ_DevMode, Len(typ_DevMode))
    '        hResult = GlobalUnlock(typ_PrintDlg.hDevMode)
    '    End If
    '
    '    'Set DEVNAMES structure from TPRINTDLG
    '    m_lng_DevNames = GlobalLock(typ_PrintDlg.hDevNames)
    '    If m_lng_DevNames > 0 Then
    '        Call CopyMemory(ByVal m_lng_DevNames, typ_DevNames, Len(typ_DevNames))
    '        hResult = GlobalUnlock(typ_PrintDlg.hDevNames)
    '    End If
    '
    '    'Set DEVMODE structure from TPRINTDLG
    '    m_lng_DevMode = GlobalLock(typ_PrintSetupDlg.hDevMode)
    '    If m_lng_DevMode > 0 Then
    '        Call CopyMemory(ByVal m_lng_DevMode, typ_DevMode, Len(typ_DevMode))
    '        hResult = GlobalUnlock(typ_PrintSetupDlg.hDevMode)
    '    End If
    '
    '    'Set DEVNAMES structure from TPRINTDLG
    '    m_lng_DevNames = GlobalLock(typ_PrintSetupDlg.hDevNames)
    '    If m_lng_DevNames > 0 Then
    '        Call CopyMemory(ByVal m_lng_DevNames, typ_DevNames, Len(typ_DevNames))
    '        hResult = GlobalUnlock(typ_PrintSetupDlg.hDevNames)
    '    End If

    ' Show Print dialog
    If PageSetupDlg(typ_PrintSetupDlg) Then
        vbInitPageSetupDlg = True

        If Not GetPrinterStructs(typ_PrintSetupDlg.hDevMode, typ_PrintSetupDlg.hDevNames, typ_DevMode, typ_DevNames) Then
            vbInitPageSetupDlg = False
        End If

    Else
        vbInitPageSetupDlg = False
    End If
End Function

Private Function GetPrinterStructs( _
    ByRef in_lng_hDevMode As Long, _
    ByRef in_lng_hDevNames As Long, _
    ByRef in_typ_DevMode As DEVMODE, _
    ByRef in_typ_DevNames As DEVNAMES, _
    Optional in_bool_SetPrntr As Boolean = False) As Boolean

    Dim str_DevName     As String
    Dim SelPrinter      As Printer
    Dim hResult         As Long

    GetPrinterStructs = True

    'Get DEVMODE structure from TPRINTDLG
    m_lng_DevMode = GlobalLock(in_lng_hDevMode)
    Call CopyMemory(in_typ_DevMode, ByVal m_lng_DevMode, Len(in_typ_DevMode))
    hResult = GlobalUnlock(m_lng_DevMode)

    'Get DEVNAMES structure from TPRINTDLG
    m_lng_DevNames = GlobalLock(in_lng_hDevNames)
    Call CopyMemory(in_typ_DevNames, ByVal m_lng_DevNames, Len(in_typ_DevNames))
    hResult = GlobalUnlock(m_lng_DevNames)

    If in_bool_SetPrntr Then
        'Set Selected Printer
        On Error Resume Next
        str_DevName = StripNulls(BytesToStr(in_typ_DevMode.dmDeviceName))
        If UCase(Printer.DeviceName) <> UCase(str_DevName) Then
            g_int_CurrentPrinterIndex = -1
            For Each SelPrinter In Printers
                g_int_CurrentPrinterIndex = g_int_CurrentPrinterIndex + 1
                If UCase(SelPrinter.DeviceName) = UCase(str_DevName) Then
                    If Not ISetDfltPrinter.SetPrinterAsDefault(SelPrinter.DeviceName) Then GetPrinterStructs = False
                    Exit For
                End If
            Next SelPrinter
        End If

        ' Set default printer properties
        On Error Resume Next
        Printer.Print "";

        #If ShowDebugPrints = 1 Then
            Debug.Print "***************Pre-setting***************"
            Debug.Print "Printer.ColorMode:", Printer.ColorMode
            Debug.Print "Printer.Copies:", Printer.Copies
            Debug.Print "Printer.Duplex:", Printer.Duplex
            Debug.Print "Printer.Orientation:", Printer.Orientation
            Debug.Print "Printer.PaperBin:", Printer.PaperBin
            Debug.Print "Printer.PaperSize:", Printer.PaperSize
            Debug.Print "Printer.PrintQuality:", Printer.PrintQuality
            Debug.Print "***************Setting***************"
        #End If

        Printer.ColorMode = in_typ_DevMode.dmColor
        Printer.Copies = in_typ_DevMode.dmCopies
        Printer.Duplex = in_typ_DevMode.dmDuplex
        Printer.Orientation = in_typ_DevMode.dmOrientation
        Printer.PaperBin = in_typ_DevMode.dmDefaultSource
        Printer.PaperSize = in_typ_DevMode.dmPaperSize
        Printer.PrintQuality = in_typ_DevMode.dmPrintQuality

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
    End If
End Function

' Convert an ANSI string in a byte array to a VB Unicode string
Public Function BytesToStr(ab() As Byte) As String
    BytesToStr = StrConv(ab, vbUnicode)
End Function

Public Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function
