VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Shape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
' Name:         clsPrintShape.cls
' Description:  Defines the Shape Object
' Author:       Vincent Lozada
' Description:
'**********************************************************************************

Option Explicit

Public BorderColor  As OLE_COLOR
Public BorderStyle  As DrawStyleConstants    'ShapeBorderStyleConstants
Public BorderWidth  As Integer
Public Height       As Long
Public Left         As Long
Public Name         As String
Public Shape        As ShapeTypeConstants
Public Top          As Long
Public Width        As Long
Public X1           As Long
Public X2           As Long
Public Y1           As Long
Public Y2           As Long

