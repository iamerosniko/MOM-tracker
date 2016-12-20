Attribute VB_Name = "SL_Listbox"
Option Compare Database

'this will get selected in listbox and returns a string value for IN CLAUSE
Public Function GetSelectedFromListBox(ByRef lb As ListBox) As String
    'value holder
    Dim tempStr As String
   
    For Each Item In lb.ItemsSelected
        tempStr = tempStr & "'" & BypassQuote(lb.ItemData(Item)) & "'" & ","
    Next Item
    
    If Len(tempStr) > 0 Then
        GetSelectedFromListBox = Left(tempStr, Len(tempStr) - 1)
    End If
End Function


