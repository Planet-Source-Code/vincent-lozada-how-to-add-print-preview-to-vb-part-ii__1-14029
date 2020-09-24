VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
' Name:         clsFields.cls
' Description:  Class Collection of the DataField Object
' Author:       Vincent Lozada
' Description:  Standard business operations supported by this business object:
'                   Add     Adds a member to the Fields Collection object
'
'                   Clear   Removes all members from the Fields Collection object
'
'                   Item    Returns a specific member of a DataField Collection object
'                            either by position or key
'
'                   Remove  Removes a member from the Fields Collection object
'                            either by position or key
'**********************************************************************************

Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private Class Members
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private m_col_DataFields As Collection

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private Class Constructor/Deconstructor
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_col_DataFields = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_col_DataFields = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Class Properties
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Count() As Integer
Attribute Count.VB_Description = "Returns the number of members in Fields collection"
    Count = m_col_DataFields.Count
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Class Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Add(x As DataField, Optional in_str_Key As String)
Attribute Add.VB_Description = "Adds a member to the Fields Collection object"
    If in_str_Key = "" Then
        Call m_col_DataFields.Add(x)
    Else
        Call m_col_DataFields.Add(x, in_str_Key)
    End If
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all members from the Fields Collection object"
    Set m_col_DataFields = New Collection
End Sub

Public Function Item(ByVal x As Variant) As DataField
Attribute Item.VB_Description = "Returns a specific member of a Fields Collection object either by position or key"
Attribute Item.VB_UserMemId = 0
    Set Item = m_col_DataFields.Item(x)
End Function

Public Sub Remove(ByVal x As Variant)
Attribute Remove.VB_Description = "Removes a member from the Fields Collection object either by position or key"
    Call m_col_DataFields.Remove(x)
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
    Set NewEnum = m_col_DataFields.[_NewEnum]
End Function
