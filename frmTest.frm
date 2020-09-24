VERSION 5.00
Object = "*\APageBrowser.vbp"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin PrintPreview.Preview Preview1 
      Height          =   4815
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8493
      MouseZoom       =   "frmTest.frx":0000
      PaperBorderColor=   16711680
      PaperShadowOffset=   70
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   8100
      TabIndex        =   1
      Top             =   0
      Width           =   8160
      Begin VB.CommandButton cmdPrintPrint 
         Caption         =   "&Print..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdPrintNext 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2065
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdPrintPrev 
         Caption         =   "Pre&v Page"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3095
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdPrintClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4125
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "&Setup..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1035
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   7
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2715
      Left            =   120
      Picture         =   "frmTest.frx":001C
      ScaleHeight     =   2655
      ScaleWidth      =   7965
      TabIndex        =   0
      Top             =   5400
      Width           =   8025
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrintClose_Click()
    Call Unload(Me)
    End
End Sub

Private Sub cmdPrintNext_Click()
    Preview1.PreviewPage = Preview1.CurrentPage + 1
    Call UpdatePageCount
End Sub

Private Sub cmdPrintPrev_Click()
    Preview1.PreviewPage = Preview1.CurrentPage - 1
    Call UpdatePageCount
End Sub

Private Sub cmdPrintPrint_Click()
    If Preview1.ShowPrintDialog(pdPrinterSetup) Then
        Call Preview1.PrintDocument
    End If
End Sub

Private Sub cmdPrintSetup_Click()
    If Preview1.ShowPrintDialog(pdPageSetup) Then
        Call Form_Load
    End If
End Sub

Private Sub Form_Load()
    Dim obj_Page            As PrintPreview.Page
    Dim obj_Header          As PrintPreview.PageHeaderFooter
    Dim obj_Footer          As PrintPreview.PageHeaderFooter
    Dim obj_Graphic         As New PrintPreview.Graphic
    Dim obj_Shape           As New PrintPreview.Shape
    Dim obj_Label           As New PrintPreview.Label
    Dim lpXCount            As Long
    Dim lpYCount            As Long
    Dim i                   As Integer
    
    Dim lng_ThumbLeft               As Long
    Dim lng_ThumbTop                As Long
    Const clng_ThumbHGap            As Long = 150
    Const clng_ThumbVGap            As Long = 100
    Const clng_ThumbBorderMargin    As Long = 30
    Dim lng_PrevThmbColor           As Long
    Dim lng_HCenterMargin           As Long
    Dim lng_VCenterMargin           As Long
    
    Const m_PageThumbWidth As Integer = 2700 '2775 '1815
    Const m_PageThumbHeight As Integer = 1650 '1545 '1335
    Const m_PageThumbLblHt As Integer = 255
    
    Const nNumberOfImages As Integer = 25
    
    Call Preview1.StartNewDocment
    Set obj_Page = Preview1.Pages(1)
    Set obj_Header = Preview1.Header
    Set obj_Footer = Preview1.Footer
    
    'Header
    obj_Label.Left = Preview1.MarginLeft
    obj_Label.Top = Preview1.MarginTop - m_PageThumbLblHt
    obj_Label.Height = m_PageThumbLblHt
    obj_Label.Width = Preview1.PageWidth
    obj_Label.Alignment = taRightTop
    obj_Label.Caption = App.Path
    obj_Label.FontName = "Arial"
    obj_Label.FontSize = 8
    obj_Label.FontUnderline = True
    obj_Label.ForeColor = vbBlack
    obj_Label.MultiLine = False
    Call obj_Header.Labels.Add(obj_Label)
    Set obj_Label = Nothing
    
    'Footer
    obj_Label.Left = Preview1.MarginLeft
    obj_Label.Top = Preview1.PaperHeight - Preview1.MarginBottom + m_PageThumbLblHt
    obj_Label.Height = m_PageThumbLblHt
    obj_Label.Width = Preview1.PageWidth
    obj_Label.Alignment = taRightTop
    obj_Label.Caption = "Page {fn:PageNumber()}"
    obj_Label.FontName = "Arial"
    obj_Label.FontSize = 8
    obj_Label.FontUnderline = False
    obj_Label.ForeColor = vbBlack
    obj_Label.MultiLine = False
    Call obj_Footer.Labels.Add(obj_Label)
    Set obj_Label = Nothing
    
    lng_HCenterMargin = (Preview1.PageWidth \ (m_PageThumbWidth + clng_ThumbHGap))
    If lng_HCenterMargin > 0 Then
        lng_HCenterMargin = (Preview1.PageWidth - ((m_PageThumbWidth + clng_ThumbHGap) * lng_HCenterMargin - clng_ThumbHGap)) / 2
    End If
    lng_VCenterMargin = (Preview1.PageHeight \ (m_PageThumbHeight + clng_ThumbVGap + m_PageThumbLblHt))
    If lng_VCenterMargin > 0 Then
        lng_VCenterMargin = (Preview1.PageHeight - ((m_PageThumbHeight + clng_ThumbVGap + m_PageThumbLblHt) * lng_VCenterMargin - clng_ThumbVGap)) / 2
    End If
    
    For i = 1 To nNumberOfImages
        lng_ThumbLeft = Preview1.MarginLeft + lng_HCenterMargin + (m_PageThumbWidth * lpXCount) + (clng_ThumbHGap * lpXCount)
        lng_ThumbTop = Preview1.MarginTop + lng_VCenterMargin + (m_PageThumbHeight * lpYCount) + (m_PageThumbLblHt * lpYCount) + (clng_ThumbVGap * lpYCount)
        
        With obj_Graphic
            .PictureType = ptBitmap
            Set .Picture = Me.Picture1.Picture
            .Left = lng_ThumbLeft
            .Top = lng_ThumbTop
            .Width = m_PageThumbWidth
            .Height = m_PageThumbHeight
        End With
        Call obj_Page.Graphics.Add(obj_Graphic)
        Set obj_Graphic = Nothing
        
        With obj_Shape
            .Left = lng_ThumbLeft - clng_ThumbBorderMargin
            .Top = lng_ThumbTop - clng_ThumbBorderMargin
            .Width = m_PageThumbWidth + clng_ThumbBorderMargin * 2
            .Height = m_PageThumbHeight + clng_ThumbBorderMargin * 2
            .BorderColor = vbBlack
            .BorderStyle = dsSolid
            .BorderWidth = 1
            .Shape = stRectangle
        End With
        Call obj_Page.Shapes.Add(obj_Shape)
        Set obj_Shape = Nothing
        
        With obj_Label
            .Alignment = taRightTop
            .Caption = "Thumb " & CStr(i)
            .FontName = "Arial"
            .FontSize = 8
            .ForeColor = vbBlack
            .Height = m_PageThumbLblHt
            .Left = lng_ThumbLeft - clng_ThumbBorderMargin
            .MultiLine = False
            .Top = lng_ThumbTop + m_PageThumbHeight + clng_ThumbBorderMargin * 2
            .Width = m_PageThumbWidth + clng_ThumbBorderMargin * 2
        End With
        Call obj_Page.Labels.Add(obj_Label)
        Set obj_Label = Nothing
        
        lpXCount = lpXCount + 1
        
        If ((m_PageThumbWidth * lpXCount) + (clng_ThumbHGap * lpXCount) + m_PageThumbWidth) > Preview1.PageWidth Then
            lpXCount = 0
            lpYCount = lpYCount + 1
            If ((m_PageThumbHeight * lpYCount) + (m_PageThumbLblHt * lpYCount) + (clng_ThumbVGap * lpYCount) + m_PageThumbHeight) > Preview1.PageHeight And i <> nNumberOfImages Then
                Set obj_Page = Preview1.NewPage
                lpYCount = 0
            End If
        End If
    Next i
    
    Call Preview1.EndDoc
    Preview1.PreviewPage = 1
    Call UpdatePageCount
End Sub

Private Sub Form_Resize()
    Call Preview1.Move(0, Me.Picture2.Height + 50, Me.ScaleWidth, Me.ScaleHeight - Me.Picture2.Height - 50)
End Sub

Private Sub UpdatePageCount()
    If Preview1.Pages.Count = 1 Then
        cmdPrintNext.Enabled = False
        cmdPrintPrev.Enabled = False
    ElseIf Preview1.CurrentPage = 1 Then
        cmdPrintNext.Enabled = True
        cmdPrintPrev.Enabled = False
    ElseIf Preview1.CurrentPage = Preview1.Pages.Count Then
        cmdPrintNext.Enabled = False
        cmdPrintPrev.Enabled = True
    Else
        cmdPrintNext.Enabled = True
        cmdPrintPrev.Enabled = True
    End If
    Label1 = "Page " & Preview1.CurrentPage & " of " & Preview1.Pages.Count
End Sub
