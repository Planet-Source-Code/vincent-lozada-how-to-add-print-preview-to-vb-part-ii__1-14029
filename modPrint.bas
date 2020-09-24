Attribute VB_Name = "modPrint"
Option Explicit

'===============================================
' PUBLIC DECLARATIONS
'===============================================
Global g_bool_SendToPrinter     As Boolean

'===============================================
' PRIVATE DECLARATIONS
'===============================================
Private m_pic_Page              As PictureBox
Private m_int_PrevScaleMode     As Integer
Private m_sng_Ratio             As Single    'The size m_sng_Ratio between the actual page and
'the print preview object
Private m_sng_LRGap             As Single    'Size of the non-printable area on printer
Private m_sng_TBGap             As Single

'The actual paper size (8.5 x 11 normally):
Private m_sng_PgWidth           As Single
Private m_sng_PgHeight          As Single

Private Const m_int_ScaleMode   As Integer = vbTwips    'Scale Object to Printer's printable area
Private Const TWIPSPERINCH = 1440

Public Sub PrintStartDoc( _
    in_sng_PaperWidth As Single, _
    in_sng_PaperHeight As Single)

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    Dim int_PrinterScaleMode    As Integer
    Dim sng_HeightRatio         As Single
    Dim sng_WidthRatio          As Single

    ' Set the physical page size:
    m_sng_PgWidth = in_sng_PaperWidth
    m_sng_PgHeight = in_sng_PaperHeight

    ' Find the size of the non-printable area on the printer to
    ' use to offset coordinates. These formulas assume the
    ' non-printable area is centered on the page:
    int_PrinterScaleMode = Printer.ScaleMode
    Printer.ScaleMode = m_int_ScaleMode
    m_sng_LRGap = (m_sng_PgWidth - Printer.ScaleWidth) / 2
    m_sng_TBGap = (m_sng_PgHeight - Printer.ScaleHeight) / 2
    Printer.ScaleMode = int_PrinterScaleMode

    ' Initialize printer or preview object:
    If g_bool_SendToPrinter Then
        m_int_PrevScaleMode = Printer.ScaleMode
        Printer.ScaleMode = m_int_ScaleMode
        Printer.Print "";
    Else
        ' Scale Object to Printer's printable area:
        m_int_PrevScaleMode = Printer.ScaleMode
        m_pic_Page.ScaleMode = m_int_ScaleMode

        ' Compare the height and with ratios to determine the
        ' m_sng_Ratio to use and how to size the picture box:
        sng_HeightRatio = m_pic_Page.ScaleHeight / m_sng_PgHeight
        sng_WidthRatio = m_pic_Page.ScaleWidth / m_sng_PgWidth

        If sng_HeightRatio < sng_WidthRatio Then
            m_sng_Ratio = sng_HeightRatio
        Else
            m_sng_Ratio = sng_WidthRatio
        End If

        ' Set default properties of picture box to match printer
        ' There are many that you could add here:
        m_pic_Page.Scale (0, 0)-(m_sng_PgWidth, m_sng_PgHeight)
        On Error Resume Next    'Printer font might not exist
        m_pic_Page.FontName = Printer.FontName
        If Err Then Call Err.Clear
        m_pic_Page.FontSize = Printer.FontSize * m_sng_Ratio
        m_pic_Page.ForeColor = Printer.ForeColor
        Call m_pic_Page.Cls
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintStartDoc", Err.Description)
End Sub

Public Sub PrintCurrentX(in_sng_Val As Single)
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.CurrentX = in_sng_Val - m_sng_LRGap
    Else
        m_pic_Page.CurrentX = in_sng_Val
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintCurrentX", Err.Description)
End Sub

Public Sub PrintCurrentY(in_sng_Val As Single)
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.CurrentY = in_sng_Val - m_sng_TBGap
    Else
        m_pic_Page.CurrentY = in_sng_Val
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintCurrentY", Err.Description)
End Sub

Public Sub PrintDrawWidth(in_lng_Val As Long)
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.DrawWidth = (ConvertPixelsToTwipsX(in_lng_Val) / 10) * in_lng_Val
    Else
        m_pic_Page.DrawWidth = m_pic_Page.ScaleX(in_lng_Val, vbPixels, vbPoints)
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintDrawWidth", Err.Description)
End Sub

Public Sub PrintDrawStyle(in_int_Val As Integer)
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.DrawStyle = in_int_Val
    Else
        m_pic_Page.DrawStyle = in_int_Val
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintDrawStyle", Err.Description)
End Sub

Public Sub PrintFontBold(in_bool_Val As Boolean)
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.FontBold = in_bool_Val
    Else
        m_pic_Page.FontBold = in_bool_Val
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintFontBold", Err.Description)
End Sub

Public Sub PrintFontItalic(in_bool_Val As Boolean)
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.FontItalic = in_bool_Val
    Else
        m_pic_Page.FontItalic = in_bool_Val
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintFontItalic", Err.Description)
End Sub

Public Sub PrintFontName(in_str_Val As String)
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.FontName = in_str_Val
    Else
        m_pic_Page.FontName = in_str_Val
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintFontName", Err.Description)
End Sub

Public Sub PrintFontSize(in_sng_Val As Single)
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.FontSize = in_sng_Val
    Else
        ' Sized by m_sng_Ratio since Scale method does not effect FontSize:
        m_pic_Page.FontSize = in_sng_Val * m_sng_Ratio
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintFontSize", Err.Description)
End Sub

Public Sub PrintFontUnderline(in_bool_Val As Boolean)
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.FontUnderline = in_bool_Val
    Else
        m_pic_Page.FontUnderline = in_bool_Val
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintFontSize", Err.Description)
End Sub

Public Sub PrintForeColor(in_lng_Val As Long)
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.ForeColor = in_lng_Val
    Else
        m_pic_Page.ForeColor = in_lng_Val
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintForeColor", Err.Description)
End Sub

Public Sub PrintPrint( _
    in_str_String As String, _
    in_int_Alignment As Integer, _
    in_bool_MultiLine As Boolean, _
    in_sng_Width As Single, _
    in_sng_Height As Single)

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    Dim str_Temp    As String

    If g_bool_SendToPrinter Then
        If in_int_Alignment = 0 Then    'Left
            If Printer.TextWidth(in_str_String) > in_sng_Width And Not in_bool_MultiLine Then
                Printer.Print SizeString(Printer, CStr(in_str_String), CSng(in_sng_Width))

            ElseIf Printer.TextWidth(in_str_String) > in_sng_Width And in_bool_MultiLine Then
                Call PrintMultiLineString(Printer, in_str_String, in_sng_Width, in_sng_Height, in_int_Alignment)

            Else
                Printer.Print in_str_String

            End If

        ElseIf in_int_Alignment = 1 Then    'Right
            If Printer.TextWidth(in_str_String) > in_sng_Width Then
                str_Temp = SizeString(Printer, CStr(in_str_String), CSng(in_sng_Width))
                Printer.CurrentX = Printer.CurrentX + (in_sng_Width - Printer.TextWidth(str_Temp))
                Printer.Print str_Temp

            Else
                Printer.CurrentX = Printer.CurrentX + (in_sng_Width - Printer.TextWidth(in_str_String))
                Printer.Print in_str_String

            End If

        ElseIf in_int_Alignment = 2 Then    'Center
            If Printer.TextWidth(in_str_String) > in_sng_Width And Not in_bool_MultiLine Then
                str_Temp = SizeString(Printer, CStr(in_str_String), CSng(in_sng_Width))
                Printer.CurrentX = Printer.CurrentX + ((in_sng_Width - Printer.TextWidth(str_Temp)) / 2)
                Printer.Print str_Temp

            ElseIf Printer.TextWidth(in_str_String) > in_sng_Width And in_bool_MultiLine Then
                Call PrintMultiLineString(Printer, in_str_String, in_sng_Width, in_sng_Height, in_int_Alignment)

            Else
                Printer.CurrentX = Printer.CurrentX + ((in_sng_Width - Printer.TextWidth(in_str_String)) / 2)
                Printer.Print in_str_String

            End If
        Else
            'ERROR
        End If
    Else
        If in_int_Alignment = 0 Then    'Left
            If m_pic_Page.TextWidth(in_str_String) > in_sng_Width And Not in_bool_MultiLine Then
                m_pic_Page.Print SizeString(m_pic_Page, CStr(in_str_String), CSng(in_sng_Width))

            ElseIf m_pic_Page.TextWidth(in_str_String) > in_sng_Width And in_bool_MultiLine Then
                Call PrintMultiLineString(m_pic_Page, in_str_String, in_sng_Width, in_sng_Height, in_int_Alignment)

            Else
                m_pic_Page.Print in_str_String

            End If

        ElseIf in_int_Alignment = 1 Then    'Right
            If m_pic_Page.TextWidth(in_str_String) > in_sng_Width Then
                str_Temp = SizeString(m_pic_Page, CStr(in_str_String), CSng(in_sng_Width))
                m_pic_Page.CurrentX = m_pic_Page.CurrentX + (in_sng_Width - m_pic_Page.TextWidth(str_Temp))
                m_pic_Page.Print str_Temp

            Else
                m_pic_Page.CurrentX = m_pic_Page.CurrentX + (in_sng_Width - m_pic_Page.TextWidth(in_str_String))
                m_pic_Page.Print in_str_String

            End If

        ElseIf in_int_Alignment = 2 Then    'Center
            If m_pic_Page.TextWidth(in_str_String) > in_sng_Width And Not in_bool_MultiLine Then
                str_Temp = SizeString(m_pic_Page, CStr(in_str_String), CSng(in_sng_Width))
                m_pic_Page.CurrentX = m_pic_Page.CurrentX + ((in_sng_Width - m_pic_Page.TextWidth(str_Temp)) / 2)
                m_pic_Page.Print str_Temp

            ElseIf m_pic_Page.TextWidth(in_str_String) > in_sng_Width And in_bool_MultiLine Then
                Call PrintMultiLineString(m_pic_Page, in_str_String, in_sng_Width, in_sng_Height, in_int_Alignment)

            Else
                m_pic_Page.CurrentX = m_pic_Page.CurrentX + ((in_sng_Width - m_pic_Page.TextWidth(in_str_String)) / 2)
                m_pic_Page.Print in_str_String

            End If

        Else
            'ERROR
        End If
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintPrint", Err.Description)
End Sub

Public Sub PrintLine( _
    in_sng_X1 As Single, _
    in_sng_Y1 As Single, _
    in_sng_X2 As Single, _
    in_sng_Y2 As Single)

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.Line (in_sng_X1 - m_sng_LRGap, in_sng_Y1 - m_sng_TBGap)-(in_sng_X2 - m_sng_LRGap, in_sng_Y2 - m_sng_TBGap)
    Else
        m_pic_Page.Line (in_sng_X1, in_sng_Y1)-(in_sng_X2, in_sng_Y2)
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintLine", Err.Description)
End Sub

Public Sub PrintBox( _
    in_sng_Left As Single, _
    in_sng_Top As Single, _
    in_sng_Width As Single, _
    in_sng_Height As Single)

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.FillStyle = vbFSSolid
        Printer.Print ""
        Printer.FillStyle = vbFSTransparent
        Printer.Line (in_sng_Left - m_sng_LRGap, in_sng_Top - m_sng_TBGap)-(in_sng_Left + in_sng_Width - m_sng_LRGap, in_sng_Top + in_sng_Height - m_sng_TBGap), , B
    Else
        m_pic_Page.Line (in_sng_Left, in_sng_Top)-(in_sng_Left + in_sng_Width, in_sng_Top + in_sng_Height), , B
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintBox", Err.Description)
End Sub

Public Sub PrintFilledBox( _
    in_sng_Left As Single, _
    in_sng_Top As Single, _
    in_sng_Width As Single, _
    in_sng_Height As Single, _
    in_lng_color As Long)

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.Line (in_sng_Left - m_sng_LRGap, in_sng_Top - m_sng_TBGap)-(in_sng_Left + in_sng_Width - m_sng_LRGap, in_sng_Top + in_sng_Height - m_sng_TBGap), in_lng_color, BF
    Else
        m_pic_Page.Line (in_sng_Left, in_sng_Top)-(in_sng_Left + in_sng_Width, in_sng_Top + in_sng_Height), in_lng_color, BF
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintFilledBox", Err.Description)
End Sub

Public Sub PrintCircle(in_sng_Left As Single, in_sng_Top As Single, in_sng_Radius)
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.Circle (in_sng_Left - m_sng_LRGap, in_sng_Top - m_sng_TBGap), in_sng_Radius
    Else
        m_pic_Page.Circle (in_sng_Left, in_sng_Top), in_sng_Radius
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintCircle", Err.Description)
End Sub

Public Sub PrintNewPage()
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Call Printer.NewPage
    Else
        Call m_pic_Page.Cls
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintNewPage", Err.Description)
End Sub

Public Sub PrintPicture( _
    in_ctl_PicSource As Control, _
    ByVal in_lng_Left As Long, _
    ByVal in_lng_Top As Long, _
    ByVal in_lng_Width As Long, _
    ByVal in_lng_Height As Long)

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    ' Picture Box should have autoredraw = False, ScaleMode = Pixel
    ' Also can have visible=false, Autosize = true
    Dim typ_BITMAPINFO  As BITMAPINFO
    Dim lng_DesthDC     As Long
    Dim lng_hMem        As Long
    Dim lng_lpBits      As Long
    Dim lng_hResult     As Long

    ' Precaution:
    If in_lng_Left < m_sng_LRGap Or in_lng_Top < m_sng_TBGap Then Exit Sub
    If in_lng_Width < 0 Or in_lng_Height < 0 Then Exit Sub
    If in_lng_Width + in_lng_Left > m_sng_PgWidth - m_sng_LRGap Then Exit Sub
    If in_lng_Height + in_lng_Top > m_sng_PgHeight - m_sng_TBGap Then Exit Sub

    in_ctl_PicSource.ScaleMode = vbPixels
    in_ctl_PicSource.AutoRedraw = False
    in_ctl_PicSource.Visible = False
    in_ctl_PicSource.AutoSize = True

    If g_bool_SendToPrinter Then
        Printer.ScaleMode = vbPixels

        ' Calculate size in pixels:
        in_lng_Left = ((in_lng_Left - m_sng_LRGap) * 1440) / Printer.TwipsPerPixelX
        in_lng_Top = ((in_lng_Top - m_sng_TBGap) * 1440) / Printer.TwipsPerPixelY
        in_lng_Width = (in_lng_Width * 1440) / Printer.TwipsPerPixelX
        in_lng_Height = (in_lng_Height * 1440) / Printer.TwipsPerPixelY
        Printer.Print "";
        lng_DesthDC = Printer.hdc
    Else
        m_pic_Page.Scale
        m_pic_Page.ScaleMode = vbPixels

        ' Calculate size in pixels:
        in_lng_Left = ((in_lng_Left * 1440) / Screen.TwipsPerPixelX) * m_sng_Ratio
        in_lng_Top = ((in_lng_Top * 1440) / Screen.TwipsPerPixelY) * m_sng_Ratio
        in_lng_Width = ((in_lng_Width * 1440) / Screen.TwipsPerPixelX) * m_sng_Ratio
        in_lng_Height = ((in_lng_Height * 1440) / Screen.TwipsPerPixelY) * m_sng_Ratio
        lng_DesthDC = m_pic_Page.hdc
    End If

    typ_BITMAPINFO.bmiHeader.biSize = 40
    typ_BITMAPINFO.bmiHeader.biWidth = in_ctl_PicSource.ScaleWidth
    typ_BITMAPINFO.bmiHeader.biHeight = in_ctl_PicSource.ScaleHeight
    typ_BITMAPINFO.bmiHeader.biPlanes = 1
    typ_BITMAPINFO.bmiHeader.biBitCount = 8
    typ_BITMAPINFO.bmiHeader.biCompression = BI_RGB

    ' Enter the following two lines as one, single line:
    lng_hMem = GlobalAlloc(GMEM_MOVEABLE, (CLng(in_ctl_PicSource.ScaleWidth + 3) \ 4) * 4 * in_ctl_PicSource.ScaleHeight)    'DWORD ALIGNED
    lng_lpBits = GlobalLock(lng_hMem)

    ' Enter the following two lines as one, single line:
    lng_hResult = GetDIBits(in_ctl_PicSource.hdc, in_ctl_PicSource.Image, 0, in_ctl_PicSource.ScaleHeight, lng_lpBits, typ_BITMAPINFO, DIB_RGB_COLORS)
    If lng_hResult <> 0 Then
        ' Enter the following two lines as one, single line:
        lng_hResult = StretchDIBits(lng_DesthDC, in_lng_Left, in_lng_Top, in_lng_Width, in_lng_Height, 0, 0, in_ctl_PicSource.ScaleWidth, in_ctl_PicSource.ScaleHeight, lng_lpBits, typ_BITMAPINFO, DIB_RGB_COLORS, SRCCOPY)
    End If

    lng_hResult = GlobalUnlock(lng_hMem)
    lng_hResult = GlobalFree(lng_hMem)

    If g_bool_SendToPrinter Then
        Printer.ScaleMode = m_int_ScaleMode
    Else
        m_pic_Page.ScaleMode = m_int_ScaleMode
        m_pic_Page.Scale (0, 0)-(m_sng_PgWidth, m_sng_PgHeight)
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintPicture", Err.Description)
End Sub

Public Sub PrintPicture2( _
    in_pic_PicSource As Picture, _
    ByVal in_sng_Left As Single, _
    ByVal in_sng_Top As Single, _
    ByVal in_sng_Width As Single, _
    ByVal in_sng_Height)

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.ScaleMode = vbPixels

        ' Calculate size in pixels:
        in_sng_Left = (in_sng_Left - m_sng_LRGap) / Printer.TwipsPerPixelX
        in_sng_Top = (in_sng_Top - m_sng_TBGap) / Printer.TwipsPerPixelY
        in_sng_Width = in_sng_Width / Printer.TwipsPerPixelX
        in_sng_Height = in_sng_Height / Printer.TwipsPerPixelY
        Printer.Print "";
    End If

    If g_bool_SendToPrinter Then
        Call Printer.PaintPicture(in_pic_PicSource, in_sng_Left, in_sng_Top, in_sng_Width, in_sng_Height, , , , , vbSrcCopy)
        Printer.ScaleMode = m_int_ScaleMode
    Else
        On Error Resume Next
        Call m_pic_Page.PaintPicture(in_pic_PicSource, in_sng_Left, in_sng_Top, in_sng_Width, in_sng_Height, , , , , vbSrcCopy)
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintPicture2", Err.Description)
End Sub

Public Sub PrintEndDoc()
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    If g_bool_SendToPrinter Then
        Printer.EndDoc
        Printer.ScaleMode = m_int_PrevScaleMode
    Else
        m_pic_Page.ScaleMode = m_int_PrevScaleMode
    End If

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintEndDoc", Err.Description)
End Sub

Public Function PrintByPage(in_obj_Page As Page, Optional EndDoc As Boolean = True) As Boolean
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    Dim i As Integer

    With in_obj_Page
        'Set local page object
        Set m_pic_Page = .PrinterObject
        'Initiate Page
        Call PrintStartDoc(.PaperWidth, .PaperHeight)
        'Header
        With in_obj_Page.Header
            Call PrintShapes(.Shapes)
            Call PrintGraphics(.Graphics)
            Call PrintLabels(.Labels)
        End With
        'Document
        Call PrintShapes(.Shapes)
        Call PrintGraphics(.Graphics)
        Call PrintFields(.Fields)
        Call PrintLabels(.Labels)
        'Footer
        With in_obj_Page.Footer
            Call PrintShapes(.Shapes)
            Call PrintGraphics(.Graphics)
            Call PrintLabels(.Labels, in_obj_Page)
        End With
        'Finalize Page
        If EndDoc Then Call PrintEndDoc
    End With

    PrintByPage = True

    Exit Function

Err_Handler:
    PrintByPage = False
    Call Err.Raise(Err.Number, App.Title & " - PrintByPage", Err.Description)
End Function

Private Sub PrintShapes(in_obj_Shapes As Shapes)
    Dim i As Integer

    For i = 1 To in_obj_Shapes.Count
        With in_obj_Shapes(i)
            Call PrintForeColor(.BorderColor)
            Call PrintDrawStyle(.BorderStyle)
            Call PrintDrawWidth(.BorderWidth)

            Select Case .Shape
                Case stRectangle, stSquare
                    Call PrintBox(.Left, .Top, .Width, .Height)
                Case stOval
                Case stCircle
                Case stRoundedRectangle
                Case stRoundedSquare
                Case stLine
                    Call PrintLine(.X1, .Y1, .X2, .Y2)
            End Select
        End With
    Next i
End Sub

Private Sub PrintGraphics(in_obj_Graphics As Graphics)
    Dim i As Integer

    For i = 1 To in_obj_Graphics.Count
        With in_obj_Graphics(i)
            Call PrintPicture2(.Picture, .Left, .Top, .Width, .Height)
        End With
    Next i
End Sub

Private Sub PrintFields(in_obj_Fields As DataFields)
    Dim i As Integer

    For i = 1 To in_obj_Fields.Count
        With in_obj_Fields(i)
            Call PrintFontName(.FontName)
            Call PrintFontSize(.FontSize)
            Call PrintFontBold(.FontBold)
            Call PrintFontItalic(.FontItalic)
            Call PrintForeColor(.ForeColor)
            Call PrintCurrentX(.Left)
            Call PrintCurrentY(.Top)
            Call PrintPrint(.DataMember, .Alignment, .MultiLine, .Width, .Height)
        End With
    Next i
End Sub

Private Sub PrintLabels(in_obj_Labels As Labels, Optional in_obj_Page As Page = Nothing)
    Dim str_Temp    As String
    Dim i           As Integer

    For i = 1 To in_obj_Labels.Count
        With in_obj_Labels(i)
            Call PrintFontName(.FontName)
            Call PrintFontSize(.FontSize)
            Call PrintFontBold(.FontBold)
            Call PrintFontItalic(.FontItalic)
            Call PrintFontUnderline(.FontUnderline)
            Call PrintForeColor(.ForeColor)
            Call PrintCurrentX(.Left)
            Call PrintCurrentY(.Top)

            If InStr(.Caption, "{fn:PageNumber()}") <> 0 And Not in_obj_Page Is Nothing Then
                str_Temp = Mid(.Caption, 1, InStr(.Caption, "{fn:PageNumber()}") - 1)
                str_Temp = str_Temp & in_obj_Page.PageNumber & Mid(.Caption, InStr(.Caption, "{fn:PageNumber()}") + Len("{fn:PageNumber()}"))
                Call PrintPrint(str_Temp, .Alignment, .MultiLine, .Width, .Height)

            Else
                Call PrintPrint(.Caption, .Alignment, .MultiLine, .Width, .Height)

            End If
        End With
    Next i
End Sub

Private Sub PrintMultiLineString( _
    in_obj_m_pic_Page As Object, _
    in_str_String As String, _
    in_sng_Width As Single, _
    in_sng_Height As Single, _
    in_int_Alignment As Integer)

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    Dim str_ary_Words()     As String
    Dim str_Chr             As String
    Dim str_Word            As String
    Dim str_Line            As String
    Dim i                   As Integer
    Dim bool_Submit         As Boolean
    Dim bool_Space          As Boolean
    Dim sng_Height          As Single
    Dim sng_CurrentX        As Single
    Dim sng_CurrentY        As Single
    Dim int_Line            As Integer
    Dim sng_TextHeight      As Single
    Dim bool_End            As Boolean

    ReDim str_ary_Words(0)
    i = 1

    Do While i <= Len(in_str_String)
        str_Chr = Mid(in_str_String, i, 1)

        If str_Chr = Chr(32) Then bool_Space = True Else bool_Space = False

        If bool_Submit <> bool_Space And str_Word <> "" Then
            bool_Submit = bool_Space
            ReDim Preserve str_ary_Words(UBound(str_ary_Words()) + 1)
            str_ary_Words(UBound(str_ary_Words())) = str_Word
            str_Word = str_Chr

        ElseIf i = Len(in_str_String) Then
            str_Word = str_Word & str_Chr
            ReDim Preserve str_ary_Words(UBound(str_ary_Words()) + 1)
            str_ary_Words(UBound(str_ary_Words())) = str_Word

        Else
            str_Word = str_Word & str_Chr

        End If

        i = i + 1
    Loop

    sng_CurrentX = in_obj_m_pic_Page.CurrentX
    sng_CurrentY = in_obj_m_pic_Page.CurrentY
    sng_TextHeight = in_obj_m_pic_Page.TextHeight("Vincent")
    int_Line = 1

    For i = 1 To UBound(str_ary_Words())
        str_Line = str_Line + str_ary_Words(i)
        If i = UBound(str_ary_Words()) And Not bool_End Then bool_End = True

        If in_obj_m_pic_Page.TextWidth(str_Line) > in_sng_Width And Not bool_End Then
            str_Line = Mid(str_Line, 1, Len(str_Line) - Len(str_ary_Words(i)))

            If str_Line <> "" Then
                If in_int_Alignment = 0 Then    'Left
                    in_obj_m_pic_Page.CurrentX = sng_CurrentX
                    in_obj_m_pic_Page.Print str_Line

                ElseIf in_int_Alignment = 1 Then    'Right
                    in_obj_m_pic_Page.CurrentX = sng_CurrentX + (in_sng_Width - in_obj_m_pic_Page.TextWidth(str_Line))
                    in_obj_m_pic_Page.Print str_Line

                ElseIf in_int_Alignment = 2 Then    'Center
                    in_obj_m_pic_Page.CurrentX = sng_CurrentX + ((in_sng_Width - in_obj_m_pic_Page.TextWidth(str_Line)) / 2)
                    in_obj_m_pic_Page.Print str_Line

                End If

                in_obj_m_pic_Page.CurrentY = sng_CurrentY + sng_TextHeight * int_Line
                sng_Height = sng_Height + sng_TextHeight

                If (sng_Height + sng_TextHeight) > in_sng_Height Then Exit For

                int_Line = int_Line + 1
                str_Line = ""

                If i = UBound(str_ary_Words()) Then bool_End = True
                i = i - 1
            Else
                str_Line = str_ary_Words(i)

                Do While (str_Line <> "" Or sng_Height < in_sng_Height)
                    str_Word = SizeString(in_obj_m_pic_Page, str_Line, in_sng_Width)

                    If in_int_Alignment = 0 Then    'Left
                        in_obj_m_pic_Page.CurrentX = sng_CurrentX
                        in_obj_m_pic_Page.Print str_Word

                    ElseIf in_int_Alignment = 1 Then    'Right
                        in_obj_m_pic_Page.CurrentX = sng_CurrentX + (in_sng_Width - in_obj_m_pic_Page.TextWidth(str_Word))
                        in_obj_m_pic_Page.Print str_Word

                    ElseIf in_int_Alignment = 2 Then    'Center
                        in_obj_m_pic_Page.CurrentX = sng_CurrentX + ((in_sng_Width - in_obj_m_pic_Page.TextWidth(str_Word)) / 2)
                        in_obj_m_pic_Page.Print str_Word

                    End If

                    in_obj_m_pic_Page.CurrentY = sng_CurrentY + sng_TextHeight * int_Line
                    sng_Height = sng_Height + sng_TextHeight

                    If (sng_Height + sng_TextHeight) > in_sng_Height Then Exit For

                    str_Line = Mid(str_Line, Len(str_Word) + 1)
                    int_Line = int_Line + 1

                    If in_obj_m_pic_Page.TextWidth(str_Line) < in_sng_Width Then Exit Do
                Loop

            End If

        ElseIf i = UBound(str_ary_Words()) And in_obj_m_pic_Page.TextWidth(str_Line) < in_sng_Width Then

            If in_int_Alignment = 0 Then    'Left
                in_obj_m_pic_Page.CurrentX = sng_CurrentX
                in_obj_m_pic_Page.Print str_Line

            ElseIf in_int_Alignment = 1 Then    'Right
                in_obj_m_pic_Page.CurrentX = sng_CurrentX + (in_sng_Width - in_obj_m_pic_Page.TextWidth(str_Line))
                in_obj_m_pic_Page.Print str_Line

            ElseIf in_int_Alignment = 2 Then    'Center
                in_obj_m_pic_Page.CurrentX = sng_CurrentX + ((in_sng_Width - in_obj_m_pic_Page.TextWidth(str_Line)) / 2)
                in_obj_m_pic_Page.Print str_Line

            End If

        ElseIf i = UBound(str_ary_Words()) And in_obj_m_pic_Page.TextWidth(str_Line) > in_sng_Width Then
            str_Line = Mid(str_Line, 1, Len(str_Line) - Len(str_ary_Words(i)))

            If str_Line <> "" Then

                If in_int_Alignment = 0 Then    'Left
                    in_obj_m_pic_Page.CurrentX = sng_CurrentX
                    in_obj_m_pic_Page.Print str_Line

                ElseIf in_int_Alignment = 1 Then    'Right
                    in_obj_m_pic_Page.CurrentX = sng_CurrentX + (in_sng_Width - in_obj_m_pic_Page.TextWidth(str_Line))
                    in_obj_m_pic_Page.Print str_Line

                ElseIf in_int_Alignment = 2 Then    'Center
                    in_obj_m_pic_Page.CurrentX = sng_CurrentX + ((in_sng_Width - in_obj_m_pic_Page.TextWidth(str_Line)) / 2)
                    in_obj_m_pic_Page.Print str_Line

                End If

                in_obj_m_pic_Page.CurrentY = sng_CurrentY + sng_TextHeight * int_Line
                sng_Height = sng_Height + sng_TextHeight

                If (sng_Height + sng_TextHeight) > in_sng_Height Then Exit For

                int_Line = int_Line + 1
                str_Line = ""

                If i = UBound(str_ary_Words()) Then bool_End = True
                i = i - 1
            Else
                str_Line = str_ary_Words(i)

                Do While (str_Line <> "" Or sng_Height < in_sng_Height)
                    str_Word = SizeString(in_obj_m_pic_Page, str_Line, in_sng_Width)

                    If in_int_Alignment = 0 Then    'Left
                        in_obj_m_pic_Page.CurrentX = sng_CurrentX
                        in_obj_m_pic_Page.Print str_Word

                    ElseIf in_int_Alignment = 1 Then    'Right
                        in_obj_m_pic_Page.CurrentX = sng_CurrentX + (in_sng_Width - in_obj_m_pic_Page.TextWidth(str_Word))
                        in_obj_m_pic_Page.Print str_Word

                    ElseIf in_int_Alignment = 2 Then    'Center
                        in_obj_m_pic_Page.CurrentX = sng_CurrentX + ((in_sng_Width - in_obj_m_pic_Page.TextWidth(str_Word)) / 2)
                        in_obj_m_pic_Page.Print str_Word

                    End If

                    in_obj_m_pic_Page.CurrentY = sng_CurrentY + sng_TextHeight * int_Line
                    sng_Height = sng_Height + sng_TextHeight
                    str_Line = Mid(str_Line, Len(str_Word) + 1)
                    int_Line = int_Line + 1

                Loop
            End If
        End If
    Next i

    Exit Sub

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - PrintMultiLineString", Err.Description)
End Sub

Private Function SizeString( _
    in_obj_m_pic_Page As Object, _
    in_str_String As String, _
    in_sng_Width As Single _
    ) As String

    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    Dim str_Temp    As String
    Dim i           As Integer: i = 1

    If in_obj_m_pic_Page.TextWidth(in_str_String) > in_sng_Width Then
        Do While in_obj_m_pic_Page.TextWidth(str_Temp & " ") < in_sng_Width
            str_Temp = str_Temp & Mid(in_str_String, i, 1)
            i = i + 1
        Loop
        SizeString = str_Temp
    Else
        SizeString = in_str_String
    End If

    Exit Function

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - SizeString", Err.Description)
End Function

Private Function ConvertPixelsToTwipsX(x As Long) As Long
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    Dim lng_hDC As Long
    Dim lng_hWnd As Long
    Dim lng_RetVal As Long
    Dim lng_XPIXELSPERINCH As Long

    '' Retrieve the current number of pixels per inch, which is
    '' resolution-dependent.
    lng_hDC = GetDC(0)
    lng_XPIXELSPERINCH = GetDeviceCaps(lng_hDC, LOGPIXELSX)
    lng_RetVal = ReleaseDC(0, lng_hDC)

    '' Compute and return the measurements in twips.
    ConvertPixelsToTwipsX = (x / lng_XPIXELSPERINCH) * TWIPSPERINCH

    Exit Function

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - ConvertPixelsToTwipsX", Err.Description)
End Function

Private Function ConvertPixelsToTwipsY(Y As Long) As Long
    #If RunErrHandler = 1 Then
        On Error GoTo Err_Handler
    #Else
        On Error GoTo 0
    #End If

    Dim lng_hDC As Long
    Dim lng_hWnd As Long
    Dim lng_RetVal As Long
    Dim lng_YPIXELSPERINCH As Long

    ' Retrieve the current number of pixels per inch, which is
    ' resolution-dependent.
    lng_hDC = GetDC(0)
    lng_YPIXELSPERINCH = GetDeviceCaps(lng_hDC, LOGPIXELSY)
    lng_RetVal = ReleaseDC(0, lng_hDC)

    ' Compute and return the measurements in twips.
    ConvertPixelsToTwipsY = (Y / lng_YPIXELSPERINCH) * TWIPSPERINCH

    Exit Function

Err_Handler:
    Call Err.Raise(Err.Number, App.Title & " - ConvertPixelsToTwipsY", Err.Description)
End Function



















