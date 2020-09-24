VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Graphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
' Name:         clsGraphics.cls
' Description:  Class Collection of the Graphic Object
' Author:       Vincent Lozada
' Description:  Standard business operations supported by this business object:
'                   Add     Adds a member to the Graphics Collection object
'
'                   Clear   Removes all members from the Graphics Collection object
'
'                   Item    Returns a specific member of a Graphic Collection object
'                            either by position or key
'
'                   Remove  Removes a member from the Graphics Collection object
'                            either by position or key
'**********************************************************************************

Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private Class Members
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private m_col_Graphics As Collection

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private Class Constructor/Deconstructor
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_col_Graphics = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_col_Graphics = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Class Properties
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Count() As Integer
Attribute Count.VB_Description = "Returns the number of members in Graphics collection"
    Count = m_col_Graphics.Count
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Class Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Add(x As Graphic, Optional in_str_Key As String)
Attribute Add.VB_Description = "Adds a member to the Graphics Collection object"
    If in_str_Key = "" Then
        Call m_col_Graphics.Add(x)
    Else
        Call m_col_Graphics.Add(x, in_str_Key)
    End If
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all members from the Graphics Collection object"
    Set m_col_Graphics = New Collection
End Sub

Public Function Item(ByVal x As Variant) As Graphic
Attribute Item.VB_Description = "Returns a specific member of a Graphics Collection object either by position or key"
Attribute Item.VB_UserMemId = 0
    Set Item = m_col_Graphics.Item(x)
End Function

Public Sub Remove(ByVal x As Variant)
Attribute Remove.VB_Description = "Removes a member from the Graphics Collection object either by position or key"
    Call m_col_Graphics.Remove(x)
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'An enumerator is a small object that knows how to
    'iterate through the items in a collection

    'The square brackets around the Collection object's
    '_NewEnum method are necessary because of the leading
    'underscore in the method name. This leading underscore
    'is a convention indicating that the method is hidden
    'in the type library.
    Set NewEnum = m_col_Graphics.[_NewEnum]
End Function

