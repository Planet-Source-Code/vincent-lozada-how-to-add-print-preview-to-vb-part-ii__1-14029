Attribute VB_Name = "modStackLog"
#If RunStackLogger = 1 Then
    Option Explicit

    Public Sub LogStackItem(in_str_Name As String)
        Dim int_StackLog As Integer
        int_StackLog = FreeFile()
        Open App.Path & "\" & App.Title & ".log" For Append As #int_StackLog
        Print #int_StackLog, Timer & " - " & in_str_Name
        Close #int_StackLog
    End Sub
#End If
