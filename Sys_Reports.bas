Attribute VB_Name = "Sys_Reports"
Option Compare Database

Public Function ProjectTrackerSQL() As String
    ProjectTrackerSQL = "Select * from BT_ProjectDetails "
End Function

Public Function ProjectNotSelect() As String
    ProjectNotSelect = " SELECT BT_Projects.ID, BT_Projects.[Project Name], BT_Projects.[Show To Reports] from BT_Projects WHERE BT_Projects.[Show To Reports] = false"
End Function

Public Function ProjectSelect() As String
    ProjectSelect = " SELECT BT_Projects.ID, BT_Projects.[Project Name], BT_Projects.[Show To Reports] from BT_Projects WHERE BT_Projects.[Show To Reports] = true"
End Function

Public Function ResetReports()
    Dim con As New DL_DA_Generic
    con.ManipulateData "UPDATE BT_Projects SET BT_Projects.[Show To Reports] = false"
End Function

Public Function GetQueryStr(ByVal QueryName As String) As String
    GetQueryStr = CurrentDb.QueryDefs(QueryName).SQL
End Function

Public Sub BuildQuery(ByVal QueryName As String, ByRef SQLCMD As String)
    CurrentDb.QueryDefs(QueryName).SQL = SQLCMD
End Sub

Public Function getWorkingDays(dateFrom, dateTo)
    Dim temp As Date
    Dim ctr As Integer
    
    If IsNull(dateFrom) Or IsNull(dateTo) Then
        getWorkingDays = ""
        Exit Function
    End If
    
    For temp = CDate(dateFrom) To CDate(dateTo)
        If Not (Weekday(temp) = 7 Or Weekday(temp) = 1) Then
            ctr = ctr + 1
        End If
    Next temp
    
    getWorkingDays = ctr
End Function

''comment
Public Function GetSelectedFromListBox(ByRef lb As ListBox) As String
    'value holder
    Dim tempstr As String
    
    For Each Item In lb.ItemsSelected
        
        tempstr = tempstr & lb.ItemData(Item) & ","
        MsgBox tempstr
    Next Item
    
    If Len(tempstr) > 0 Then
        GetSelectedFromListBox = Left(tempstr, Len(tempstr) - 1)
    End If
End Function

