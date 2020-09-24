VERSION 5.00
Begin VB.Form frmAbort 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Printing..."
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   330
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblDeviceNameCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "on the HP Laser Jet on LPT1:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblFileName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "My Document"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblPageCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Now Printing 1 of 1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmAbort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_enum_AbortWindowPosition      As AbortWindowPositionConstants

Public Property Get AbortWindowCaption() As String
    AbortWindowCaption = Me.Caption
End Property
Public Property Let AbortWindowCaption(x As String)
    Me.Caption = x
End Property

Public Property Get AbortButtonCaption() As String
    AbortButtonCaption = cmdCancel.Caption
End Property
Public Property Let AbortButtonCaption(x As String)
    cmdCancel.Caption = x
End Property

Public Property Get AbortDeviceNameCaption() As String
    AbortDeviceNameCaption = lblDeviceNameCaption.Caption
End Property
Public Property Let AbortDeviceNameCaption(x As String)
    lblDeviceNameCaption.Caption = x
End Property

Public Property Get AbortFileNameCaption() As String
    AbortFileNameCaption = lblFileName.Caption
End Property
Public Property Let AbortFileNameCaption(x As String)
    lblFileName.Caption = x
End Property

Public Property Get AbortPageCaption() As String
    AbortPageCaption = lblPageCaption.Caption
End Property
Public Property Let AbortPageCaption(x As String)
    lblPageCaption.Caption = x
End Property

Public Property Get AbortWindowPosition() As AbortWindowPositionConstants
    AbortWindowPosition = m_enum_AbortWindowPosition
End Property
Public Property Let AbortWindowPosition(x As AbortWindowPositionConstants)
    m_enum_AbortWindowPosition = x
End Property

Private Sub Form_Load()
    'Initialize Controls
    lblDeviceNameCaption.BorderStyle = 0
    lblFileName.BorderStyle = 0
    lblPageCaption.BorderStyle = 0

    Select Case m_enum_AbortWindowPosition
        Case awpAppWindow
        Case awpScreenCenter
    End Select
End Sub


















