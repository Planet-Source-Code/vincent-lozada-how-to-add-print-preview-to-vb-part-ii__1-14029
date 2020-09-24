VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
' Name:         clsPrintPage.cls
' Description:  Defines the Page Object
' Author:       Vincent Lozada
' Description:
'**********************************************************************************

Option Explicit

'===============================================
' PUBLIC DECLARATIONS
'===============================================
Public Fields           As DataFields
Public Graphics         As Graphics
Public Labels           As Labels
Public Shapes           As Shapes
Public Header           As PageHeaderFooter
Public Footer           As PageHeaderFooter

'===============================================
' PRIVATE DECLARATIONS
'===============================================
'Read/Write
Private m_lng_MarginBottom      As Long
Private m_lng_MarginFooter      As Long
Private m_lng_MarginHeader      As Long
Private m_lng_MarginLeft        As Long
Private m_lng_MarginRight       As Long
Private m_lng_MarginTop         As Long
Private m_enum_Orientation      As OrientationConstants
Private m_enum_PageBorder       As PageBorderConstants
Private m_lng_PageNumber        As Long
Private m_lng_PageHeight        As Long
Private m_lng_PageWidth         As Long
Private m_lng_PaperHeight       As Long
Private m_enum_PaperSize        As PaperSizeConstants
Private m_lng_PaperWidth        As Long
Private m_enum_ShowGuides       As ShowGuidesConstants

Private m_obj_PrinterObject     As Control

'*************************************************************
' PUBLIC READ/WRITE PROPERTIES
'*************************************************************
Public Property Get DrawStyle() As DrawStyleConstants
    DrawStyle = m_obj_PrinterObject.DrawStyle
End Property
Public Property Let DrawStyle(ByVal New_DrawStyle As DrawStyleConstants)
    m_obj_PrinterObject.DrawStyle() = New_DrawStyle
End Property

Public Property Get DrawWidth() As Integer
    DrawWidth = m_obj_PrinterObject.DrawWidth
End Property
Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    m_obj_PrinterObject.DrawWidth() = New_DrawWidth
End Property

Public Property Get FillColor() As OLE_COLOR
    FillColor = m_obj_PrinterObject.FillColor
End Property
Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    m_obj_PrinterObject.FillColor() = New_FillColor
End Property

Public Property Get FillStyle() As FillStyleConstants
    FillStyle = m_obj_PrinterObject.FillStyle
End Property
Public Property Let FillStyle(ByVal New_FillStyle As FillStyleConstants)
    m_obj_PrinterObject.FillStyle() = New_FillStyle
End Property

Public Property Get Font() As Font
    Set Font = m_obj_PrinterObject.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set m_obj_PrinterObject.Font = New_Font
End Property

Public Property Get FontBold() As Boolean
    FontBold = m_obj_PrinterObject.FontBold
End Property
Public Property Let FontBold(ByVal New_FontBold As Boolean)
    m_obj_PrinterObject.FontBold() = New_FontBold
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = m_obj_PrinterObject.FontItalic
End Property
Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    m_obj_PrinterObject.FontItalic() = New_FontItalic
End Property

Public Property Get FontName() As String
    FontName = m_obj_PrinterObject.FontName
End Property
Public Property Let FontName(ByVal New_FontName As String)
    m_obj_PrinterObject.FontName() = New_FontName
End Property

Public Property Get FontSize() As Single
    FontSize = m_obj_PrinterObject.FontSize
End Property
Public Property Let FontSize(ByVal New_FontSize As Single)
    m_obj_PrinterObject.FontSize() = New_FontSize
End Property

Public Property Get FontStrikethru() As Boolean
    FontStrikethru = m_obj_PrinterObject.FontStrikethru
End Property
Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    m_obj_PrinterObject.FontStrikethru() = New_FontStrikethru
End Property

Public Property Get FontTransparent() As Boolean
    FontTransparent = m_obj_PrinterObject.FontTransparent
End Property
Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
    m_obj_PrinterObject.FontTransparent() = New_FontTransparent
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = m_obj_PrinterObject.FontUnderline
End Property
Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    m_obj_PrinterObject.FontUnderline() = New_FontUnderline
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_obj_PrinterObject.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_obj_PrinterObject.ForeColor() = New_ForeColor
End Property

Property Get MarginBottom() As Long
    MarginBottom = m_lng_MarginBottom
End Property
Property Let MarginBottom(x As Long)
    m_lng_MarginBottom = x
End Property

Property Get MarginLeft() As Long
    MarginLeft = m_lng_MarginLeft
End Property
Property Let MarginLeft(x As Long)
    m_lng_MarginLeft = x
End Property

Property Get MarginRight() As Long
    MarginRight = m_lng_MarginRight
End Property
Property Let MarginRight(x As Long)
    m_lng_MarginRight = x
End Property

Property Get MarginTop() As Long
    MarginTop = m_lng_MarginTop
End Property
Property Let MarginTop(x As Long)
    m_lng_MarginTop = x
End Property

Property Get Orientation() As OrientationConstants
    Orientation = m_enum_Orientation
End Property
Property Let Orientation(x As OrientationConstants)
    m_enum_Orientation = x
End Property

Property Get PageBorder() As PageBorderConstants
    PageBorder = m_enum_PageBorder
End Property
Property Let PageBorder(x As PageBorderConstants)
    m_enum_PageBorder = x
End Property

Property Get PageHeight() As Long
    PageHeight = m_lng_PageHeight
End Property
Property Let PageHeight(x As Long)
    m_lng_PageHeight = x
End Property

Property Get PageNumber() As Long
    PageNumber = m_lng_PageNumber
End Property
Friend Property Let PageNumber(x As Long)
    m_lng_PageNumber = x
End Property

Property Get PageWidth() As Long
    PageWidth = m_lng_PageWidth
End Property
Property Let PageWidth(x As Long)
    m_lng_PageWidth = x
End Property

Property Get PaperHeight() As Long
    PaperHeight = m_lng_PaperHeight
End Property
Property Let PaperHeight(x As Long)
    m_lng_PaperHeight = x
End Property

Property Get PaperSize() As PaperSizeConstants
    PaperSize = m_enum_PaperSize
End Property
Property Let PaperSize(x As PaperSizeConstants)
    m_enum_PaperSize = x
End Property

Property Get PaperWidth() As Long
    PaperWidth = m_lng_PaperWidth
End Property
Property Let PaperWidth(x As Long)
    m_lng_PaperWidth = x
End Property

Friend Property Get PrinterObject() As Control
    Set PrinterObject = m_obj_PrinterObject
End Property
Friend Property Let PrinterObject(x As Control)
    Set m_obj_PrinterObject = x
End Property

Property Get ShowGuides() As ShowGuidesConstants
    ShowGuides = m_enum_ShowGuides
End Property
Property Let ShowGuides(x As ShowGuidesConstants)
    m_enum_ShowGuides = x
End Property

'*************************************************************
' CLASS CONTRUCTOR/DECONSTRUCTOR
'*************************************************************
Private Sub Class_Initialize()
    Set Fields = New DataFields
    Set Graphics = New Graphics
    Set Labels = New Labels
    Set Shapes = New Shapes
End Sub

Private Sub Class_Terminate()
    Set Fields = Nothing
    Set Graphics = Nothing
    Set Labels = Nothing
    Set Shapes = Nothing
End Sub

'*************************************************************
' PUBLIC METHODS
'*************************************************************
Public Sub PrintPage()
    Call PrintByPage(Me)
End Sub

