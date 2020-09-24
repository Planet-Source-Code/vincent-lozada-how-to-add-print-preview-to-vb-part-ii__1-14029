VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Shapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
' Name:         clsShapes.cls
' Description:  Class Collection of the Shape Object
' Author:       Vincent Lozada
' Description:  Standard business operations supported by this business object:
'                   Add     Adds a member to the Shapes Collection object
'
'                   Clear   Removes all members from the Shapes Collection object
'
'                   Item    Returns a specific member of a Shape Collection object
'                            either by position or key
'
'                   Remove  Removes a member from the Shapes Collection object
'                            either by position or key
'**********************************************************************************

Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private Class Members
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private m_col_Shapes As Collection

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private Class Constructor/Deconstructor
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_col_Shapes = New Collection
End Sub

Private Sub Class_Terminate()
    Call Clear
    Set m_col_Shapes = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Class Properties
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Count() As Integer
Attribute Count.VB_Description = "Returns the number of members in Shapes collection"
    Count = m_col_Shapes.Count
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Class Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Add(x As Shape, Optional in_str_Key As String)
Attribute Add.VB_Description = "Adds a member to the Shapes Collection object"
    If in_str_Key = "" Then
        m_col_Shapes.Add x
    Else
        m_col_Shapes.Add x, in_str_Key
    End If
End Sub

Public Sub Clear()
    Set m_col_Shapes = New Collection
End Sub

Public Function Item(ByVal x As Variant) As Shape
Attribute Item.VB_Description = "Returns a specific member of a Shapes Collection object either by position or key"
Attribute Item.VB_UserMemId = 0
    Set Item = m_col_Shapes.Item(x)
End Function

Public Sub Remove(ByVal x As Variant)
Attribute Remove.VB_Description = "Removes a member from the Shapes Collection object either by position or key"
    m_col_Shapes.Remove x
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
    Set NewEnum = m_col_Shapes.[_NewEnum]
End Function

