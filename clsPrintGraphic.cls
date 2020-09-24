VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Graphic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
' Name:         clsPrintGraphic.cls
' Description:  Defines the Graphic Object
' Author:       Vincent Lozada
' Description:
'**********************************************************************************

Option Explicit

Public Enum PictureTypeConstants
    ptBitmap = 0
    ptIcon = 1
    ptCursor = 2
End Enum

Public Name         As String
Public PictureType  As PictureTypeConstants
Public Picture      As Picture    'StdPicture
Public Left         As Single
Public Top          As Single
Public Width        As Single
Public Height       As Single


