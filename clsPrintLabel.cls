VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
' Name:         clsPrintLabel.cls
' Description:  Defines the Label Object
' Author:       Vincent Lozada
' Description:
'**********************************************************************************

Option Explicit

Public Alignment        As TextAlignmentConstants
Public Caption          As String
Public Font             As New StdFont
Public FontBold         As Boolean
Public FontItalic       As Boolean
Public FontName         As String
Public FontSize         As Single
Public FontStrikethru   As Boolean
Public FontTransparent  As Boolean
Public FontUnderline    As Boolean
Public ForeColor        As OLE_COLOR
Public Height           As Long
Public IndentFirst      As Integer
Public IndentLeft       As Integer
Public IndentRight      As Integer
Public IndentTab        As Integer
Public Left             As Long
Public LineSpacing      As Integer
Public MultiLine        As Boolean
Public Name             As String
Public SpaceAfter       As Integer
Public SpaceBefore      As Integer
Public TextAngle        As Integer
Public Top              As Long
Public Width            As Long



