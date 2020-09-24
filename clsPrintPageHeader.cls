VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageHeaderFooter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
' Name:         clsPrintPageHeader.cls
' Description:  Defines the PageHeaderFooter Object
' Author:       Vincent Lozada
' Description:
'**********************************************************************************

Option Explicit

'===============================================
' PUBLIC DECLARATIONS
'===============================================
Public Height       As Single
Public Graphics     As Graphics
Public Labels       As Labels
Public Shapes       As Shapes
'===============================================

'===============================================
' Private Class Constructor/Deconstructor
'===============================================
Private Sub Class_Initialize()
    Set Graphics = New Graphics
    Set Labels = New Labels
    Set Shapes = New Shapes
End Sub

Private Sub Class_Terminate()
    Set Graphics = Nothing
    Set Labels = Nothing
    Set Shapes = Nothing
End Sub

'===============================================
' Public Class Methods
'===============================================
Public Sub Clear()
    Call Class_Initialize
End Sub
