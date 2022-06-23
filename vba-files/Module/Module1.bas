Attribute VB_Name = "Module1"
Sub Report()

    Dim filePath As String
    Let filePath = ThisWorkbook.Path

    Dim distribution_list_rng As Range
    ' use "Name Manager" to ensure that the names of each area are stored as a named range
    ' use this formula: =OFFSET(Control!$C$1,8,0,COUNTA(Control!$C:$C)-1,1)
    Set distribution_list_rng = Range("distribution_list")

    Dim b As String
    Let b = Range("filename").Value
    '^--this is to hold the file name

    Dim c As String
    Let c = "v" & Range("version").Value
    '^--this holds the version number

    Dim d As Range
    '^--use this to store each range within the distribution_list

    For Each d In distribution_list_rng

        Sheets("Report").PivotTables("report_pivot").PivotFields("State").ClearAllFilters
        '^--Income statement is a pivot tbl on the Report sheet, clear the filter
        Sheets("Report").Range("b5") = d
        '^--filter by the first value in the distribution_list
        MsgBox ("Packaging " & d)
        Sheets("Report").Select
        Sheets("Report").Copy
        '^--copy the Report sheet for a new workbook
        Application.DisplayAlerts = False

        ActiveWorkbook.SaveAs Filename:=filePath & _
            "\reports_to_distribute\" & b & " - " & d & " - " & c & ".xlsx", _
            FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        ActiveWorkbook.Save
        ActiveWorkbook.Close

    Next d

End Sub

