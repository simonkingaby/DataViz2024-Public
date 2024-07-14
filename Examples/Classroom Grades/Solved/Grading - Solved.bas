Attribute VB_Name = "Module5"
Option Explicit

Sub ImportData()
'
' ImportData Macro
'

'
    Dim book As Workbook
    Set book = Workbooks.Add
    book.Activate
    ActiveWorkbook.Queries.Add Name:="sample_class_data", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Csv.Document(File.Contents(""F:\CodeRepos\2024DataViz\DataViz2024-Public\Examples\Classroom Grades\sample_class_data.csv""),[Delimiter="","", Columns=16, Encoding=1252, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted He" & _
        "aders"",{{""Class"", type text}, {""Section"", type text}, {""Teacher"", type text}, {""Student"", type text}, {""Assignment 1"", Int64.Type}, {""Assignment 2"", Int64.Type}, {""Assignment 3"", Int64.Type}, {""Assignment 4"", Int64.Type}, {""Assignment 5"", Int64.Type}, {""Exam"", Int64.Type}, {""Assignment 6"", Int64.Type}, {""Assignment 7"", Int64.Type}, {""Assign" & _
        "ment 8"", type text}, {""Assignment 9"", type text}, {""Assignment 10"", type text}, {""Final Exam"", Int64.Type}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=sample_class_data;Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [sample_class_data]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "sample_class_data"
        .Refresh BackgroundQuery:=False
    End With
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "Raw Data"
    Sheets("Sheet1").Select
    Application.DisplayAlerts = False
    ActiveWindow.SelectedSheets.Delete
    Application.DisplayAlerts = True
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = "Summary"
    
    SplitClasses
    
End Sub

Function GetLastRow(sheet As Worksheet) As Integer
    Dim lastRow As Integer
    With sheet
          lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
    End With
    GetLastRow = lastRow
End Function

Sub CreateClassSheet(className As String)
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = className
End Sub

Function CreateSection(dataRow As Integer, previousClass As String, classRow As Integer) As Integer
    Dim class As String
    Dim section As String
    Dim teacher As String
    class = Worksheets("Raw Data").Cells(dataRow, 1).Value
    If class <> previousClass Then
        CreateClassSheet class
        classRow = 1
    End If
    
    section = Worksheets("Raw Data").Cells(dataRow, 2).Value
    teacher = Worksheets("Raw Data").Cells(dataRow, 3).Value
    ' section header
    With Worksheets(class)
        With .Cells(classRow, 1)
            .Value = section
            .Style = "40% - Accent1"
            .Style = "Heading 1"
        End With
        With .Cells(classRow, 2)
            .Value = teacher
            .Style = "40% - Accent1"
            .Style = "Heading 1"
        End With
        classRow = classRow + 1
        With .Cells(classRow, 1)
            .Value = "Student"
            .Style = "Heading 3"
        End With
        With .Cells(classRow, 2)
            .Value = "Assignment 1"
            .Style = "Heading 3"
        End With
        With .Cells(classRow, 3)
            .Value = "Assignment 2"
            .Style = "Heading 3"
        End With
        With .Cells(classRow, 4)
            .Value = "Assignment 3"
            .Style = "Heading 3"
        End With
        With .Cells(classRow, 5)
            .Value = "Assignment 4"
            .Style = "Heading 3"
        End With
        With .Cells(classRow, 6)
            .Value = "Assignment 5"
            .Style = "Heading 3"
        End With
        With .Cells(classRow, 7)
            .Value = "Mid-term"
            .Style = "Heading 3"
        End With
        With .Cells(classRow, 8)
            .Value = "Assignment 6"
            .Style = "Heading 3"
        End With
        With .Cells(classRow, 9)
            .Value = "Assignment 7"
            .Style = "Heading 3"
        End With
        With .Cells(classRow, 10)
            .Value = "Assignment 8"
            .Style = "Heading 3"
        End With
        With .Cells(classRow, 11)
            .Value = "Assignment 9"
            .Style = "Heading 3"
        End With
        With .Cells(classRow, 12)
            .Value = "Assignment 10"
            .Style = "Heading 3"
        End With
        With .Cells(classRow, 13)
            .Value = "Final Exam"
            .Style = "Heading 3"
        End With
        With .Cells(classRow, 14)
            .Value = "Final Grade"
            .Style = "Heading 3"
        End With
        .Columns("A:N").EntireColumn.AutoFit
    End With
    CreateSection = classRow + 1
End Function

Function WriteOutStudent(dataRow As Integer, classRow As Integer) As Double
    Dim class As String
    Dim raw As Worksheet
    Set raw = Worksheets("Raw Data")
    class = raw.Cells(dataRow, 1).Value
    Dim finalGrade As Double
    finalGrade = 0
    Dim i As Integer
    Dim weight As Double
    With Worksheets(class)
        weight = 1
        For i = 1 To 13
            .Cells(classRow, i).Value = raw.Cells(dataRow, i + 3).Value
            If raw.Cells(dataRow, i + 4).Value <> "N/A" Then
                If i = 6 Then
                    finalGrade = finalGrade + 0.2 * raw.Cells(dataRow, i + 4).Value
                    weight = weight - 0.2
                ElseIf i = 12 Then
                    finalGrade = finalGrade + weight * raw.Cells(dataRow, i + 4).Value
                Else
                    finalGrade = finalGrade + 0.05 * Int(raw.Cells(dataRow, i + 4).Value)
                    weight = weight - 0.05
                End If
            End If
        Next i
        .Cells(classRow, 14).Value = finalGrade
    End With
    WriteOutStudent = finalGrade
End Function

Function CloseSection(class As String, section As String, teacher As String, students As Integer, grades As Double, summaryRow As Integer, classRow As Integer) As Integer
    With Worksheets(class)
        .Cells(classRow, 1).Value = section
        .Cells(classRow, 2).Value = teacher
        .Cells(classRow, 3).Value = "Class Average"
        .Cells(classRow, 14).Value = Round(grades / students, 2)
        Range("A" & classRow & ":N" & classRow).Select
        Selection.Style = "Heading 4"
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .weight = xlMedium
        End With
    End With
    ' write to the summary sheet for the section
    With Worksheets("Summary")
        .Cells(summaryRow, 1) = class
        .Cells(summaryRow, 2) = section
        .Cells(summaryRow, 3) = teacher
        .Cells(summaryRow, 4) = Round(grades / students, 2)
    End With
    CloseSection = classRow + 2
End Function

Sub SplitClasses()
    Dim rawDataRow As Integer
    Dim lastRow As Integer
    Dim class As String
    Dim section As String
    Dim teacher As String
    Dim students As Integer
    Dim grades As Double
    Dim summaryRow As Integer
    Dim classRow As Integer
    Dim previousClass As String
        
    lastRow = GetLastRow(Worksheets("Raw Data"))
    classRow = 1
    
    'Initialize the Summary
    With Worksheets("Summary")
        With .Cells(1, 1)
            .Value = "Class"
            .Style = "40% - Accent1"
            .Style = "Heading 1"
        End With
        With .Cells(1, 2)
            .Value = "Section"
            .Style = "40% - Accent1"
            .Style = "Heading 1"
        End With
        With .Cells(1, 3)
            .Value = "Teacher"
            .Style = "40% - Accent1"
            .Style = "Heading 1"
        End With
        With .Cells(1, 4)
            .Value = "Section Average"
            .Style = "40% - Accent1"
            .Style = "Heading 1"
        End With
    End With
    summaryRow = 2
    
    ' Initialize variables from first data row
    With Worksheets("Raw Data")
        class = .Cells(2, 1).Value
        section = .Cells(2, 2).Value
        teacher = .Cells(2, 3).Value
    End With
    classRow = CreateSection(2, "", classRow)
    students = 0
    grades = 0
    
    For rawDataRow = 2 To lastRow
        If section <> Worksheets("Raw Data").Cells(rawDataRow, 2).Value Then
            ' found a new section
            ' close out the previous section
            classRow = CloseSection(class, section, teacher, students, grades, summaryRow, classRow)
            students = 0
            grades = 0
            summaryRow = summaryRow + 1
            previousClass = class
            'Initialize next section
            With Worksheets("Raw Data")
                class = .Cells(rawDataRow, 1).Value
                section = .Cells(rawDataRow, 2).Value
                teacher = .Cells(rawDataRow, 3).Value
            End With
            classRow = CreateSection(rawDataRow, previousClass, classRow)
        Else
            'write out the student data and keep going
            grades = grades + WriteOutStudent(rawDataRow, classRow)
            classRow = classRow + 1
            students = students + 1
        End If
    Next rawDataRow
    classRow = CloseSection(class, section, teacher, students, grades, summaryRow, classRow)
    
    UpdateSummarySheet summaryRow
End Sub

Sub UpdateSummarySheet(lastRow As Integer)
    Dim currRow As Integer
    Dim class As String
    Dim amount As Double
    Dim avg As Double
    Dim totalRows As Integer
    Dim classFirstRow As Integer
    
    totalRows = 0
    classFirstRow = 2
    Worksheets("Summary").Activate
    For currRow = 2 To lastRow
        With Worksheets("Summary")
            class = .Cells(currRow + totalRows, 1).Value
            amount = .Cells(currRow + totalRows, 4).Value
            ' check next row class
            If .Cells(currRow + totalRows + 1, 1).Value <> class Then
                totalRows = totalRows + 1
                Rows((currRow + totalRows) & ":" & (currRow + totalRows)).Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                .Cells(currRow + totalRows, 1).Value = "Class Average"
                With .Cells(currRow + totalRows, 4)
                    .Formula = "=Round(Subtotal(1,D" & classFirstRow & ":D" & (currRow + totalRows - 1) & "),2)"
                    .Style = "Heading 4"
                    .Select
                    With Selection.Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .weight = xlThin
                    End With
                    With Selection.Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .weight = xlThin
                    End With
                End With
                classFirstRow = currRow + totalRows + 1
            End If
        End With
    Next currRow
    Worksheets("Summary").Columns("A:D").EntireColumn.AutoFit
    Worksheets("Summary").Activate
End Sub

