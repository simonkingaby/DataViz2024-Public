Attribute VB_Name = "ClassroomGradeSolution"
Option Explicit

' To use this module, open Excel to a new blank workbook
' press Alt-F11 to go to the code windows
' find your personal workbook in the VBAProject pane
' right-click on it and choose Insert > Module
' then paste all of this code into the module window
' press Ctrl-S to save the personal workbook and the new module
' flip back to the blank workbook and run the ImportData macro
' from the Developer toolbar, choose Macros, then select ImportData and click Run
' if we're lucky, it will import the data and split it into individual sheets by class
' and calculate the section and class averages
' and display the summary of those section and class averages

Sub ImportData()
'
' ImportData Macro
'

'
    ' First add a new workbook
    Dim book As Workbook
    Set book = Workbooks.Add
    book.Activate
    ' This is the code that is recorded when you import a CSV file
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
    ' Everything above here was as recorded
    ' Now we rename Sheet2 to Raw Data 
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "Raw Data"
    ' And we delete the first sheet
    Sheets("Sheet1").Select
    ' Because deleting a sheet will cause a prompt in the middle of the macro execution,
    ' we turn off alerts to suppress this annoying behavior
    Application.DisplayAlerts = False
    ActiveWindow.SelectedSheets.Delete
    ' Then turn them back on as soon as we're done, otherwise we won't see any errors later
    Application.DisplayAlerts = True
    ' Now add the summary sheet
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = "Summary"
    
    'Now call the custom code that splits the raw data into indivudal sheets by Class
    SplitClasses

    ' And save the workbook
    book.SaveAs Filename:="Grades.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End Sub

Function GetLastRow(sheet As Worksheet) As Integer
    'This function will return the last row in a worksheet
    Dim lastRow As Integer
    With sheet
          lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
    End With
    GetLastRow = lastRow
End Function

Sub CreateClassSheet(className As String)
    ' This subroutine will create a new worksheet with the name of the Class as the tab/sheet name
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = className
End Sub

Function CreateSection(dataRow As Integer, previousClass As String, classRow As Integer) As Integer
    ' This function will create a new section in the class sheet
    ' adding the section header and a header row for the student data
    ' Parameters:
    ' dataRow - the row in the Raw Data sheet that we are currently processing
    ' previousClass - the name of the class from the previous row
    ' classRow - the row number in the class sheet that we are currently writing to
    ' Output:
    ' Each function is responsible for keeping track of the rows it uses and incrementing classRow
    ' accordingly. Then it will return the row number of the next row to write to.
    Dim class As String
    Dim section As String
    Dim teacher As String
    ' get the class, section, and teacher from the raw data row
    class = Worksheets("Raw Data").Cells(dataRow, 1).Value
    ' if it's a new class, create a new sheet for the new class
    If class <> previousClass Then
        CreateClassSheet class
        classRow = 1
    End If
    section = Worksheets("Raw Data").Cells(dataRow, 2).Value
    teacher = Worksheets("Raw Data").Cells(dataRow, 3).Value
    ' write the section header and the student data header
    With Worksheets(class)
        'section header, white text on a dark blue background and bold
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
        ' student data header, just bold, increment the row
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
        ' now we have all the column headers established, autofit the columns
        .Columns("A:N").EntireColumn.AutoFit
    End With
    'return the next classrow to write to
    CreateSection = classRow + 1
End Function

Function WriteOutStudent(dataRow As Integer, classRow As Integer) As Double
    ' This function will write out the student data to the class sheet
    ' and calculate the final grade for the student
    ' Parameters:
    ' dataRow - the row in the Raw Data sheet that we are currently processing
    ' classRow - the row number in the class sheet that we are currently writing to
    ' Output:
    ' The final grade for the student

    Dim class As String
    'since we'll refer to the raw data sheet a lot, let's set it up here
    Dim raw As Worksheet
    Set raw = Worksheets("Raw Data")
    ' grab the class name from the raw data row column 1
    class = raw.Cells(dataRow, 1).Value
    ' initialize the final grade to 0
    Dim finalGrade As Double
    finalGrade = 0
    ' i is just a loop counter
    Dim i As Integer
    ' weight is used to calculate the final exam weight, it starts at 1 and
    ' is reduced by 0.05 for each assignment grade that is present
    ' and by .2 for the midterm grade, any leftover is the weight of the final exam grade
    Dim weight As Double
    ' since we're working on the class sheet, we'll use the With statement to make it easier
    With Worksheets(class)
        ' initialize the final exam weight to 1
        weight = 1
        'loop through the individual grades for this student
        For i = 1 To 13
            ' first, write the raw data column corresponding to i
            ' into the student's row in the class sheet
            ' note, the first column is their name
            ' the rest are grades, with up to 10 assignments, a midterm, and a final exam
            .Cells(classRow, i).Value = raw.Cells(dataRow, i + 3).Value
            ' now, if the grade is not "N/A" then we'll add this assignment or exam to the final grade           
            If raw.Cells(dataRow, i + 4).Value <> "N/A" Then
                ' if it's the midterm, we'll add 20% of the grade to the final grade
                If i = 6 Then
                    finalGrade = finalGrade + 0.2 * raw.Cells(dataRow, i + 4).Value
                    weight = weight - 0.2
                ' if it's the final exam, we'll add the remaining weight to the final grade
                ElseIf i = 12 Then
                    finalGrade = finalGrade + weight * raw.Cells(dataRow, i + 4).Value
                Else
                    ' otherwise, we'll add 5% of the assignment grade to the final grade
                    finalGrade = finalGrade + 0.05 * Int(raw.Cells(dataRow, i + 4).Value)
                    weight = weight - 0.05
                End If
            End If
        Next i
        ' write out the final grade
        .Cells(classRow, 14).Value = finalGrade
    End With
    ' return the final grade to the calling function for use in calculating the section average
    WriteOutStudent = finalGrade
End Function

Function CloseSection(class As String, section As String, teacher As String, students As Integer, grades As Double, summaryRow As Integer, classRow As Integer) As Integer
    ' This function will close out the current section
    ' by writing the section average to the class sheet
    ' and adding a row to the summary sheet
    ' Parameters:
    ' class - the name of the class
    ' section - the name of the section
    ' teacher - the name of the teacher
    ' students - the number of students in the section
    ' grades - the total grades for the section
    ' summaryRow - the row number in the summary sheet that we are currently writing to
    ' classRow - the row number in the class sheet that we are currently writing to
    ' Output:
    ' The next row to write to in the class sheet
    
    ' write out the section average, again using the With statement to make it less typing
    With Worksheets(class)
        .Cells(classRow, 1).Value = section
        .Cells(classRow, 2).Value = teacher
        .Cells(classRow, 3).Value = "Section Average"
        .Cells(classRow, 14).Value = Round(grades / students, 2)
        ' format the section average row so it stands out a bit
        ' note: formatting is so much easier when you record it first and then modify the recorded code
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
    ' write a new row to the summary sheet for the section
    With Worksheets("Summary")
        .Cells(summaryRow, 1) = class
        .Cells(summaryRow, 2) = section
        .Cells(summaryRow, 3) = teacher
        .Cells(summaryRow, 4) = Round(grades / students, 2)
    End With
    ' add two so we get a blank row between sections
    CloseSection = classRow + 2
End Function

Sub SplitClasses()
    ' This subroutine will split the raw data into individual sheets by class
    ' and calculate the section averages and class averages
    ' It will also create a summary sheet with the class averages

    ' Initialize variables
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
        
    ' Get the last row of the raw data so we have the range to loop through
    lastRow = GetLastRow(Worksheets("Raw Data"))
    ' Initialize the classRow to 1 so we know where to start writing on the class sheet
    classRow = 1
    
    'Initialize the Column headers on the Summary sheet
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
    ' increment the summary row so we know where to write the first section average
    summaryRow = 2
    
    ' Initialize some variables from first data row
    With Worksheets("Raw Data")
        class = .Cells(2, 1).Value
        section = .Cells(2, 2).Value
        teacher = .Cells(2, 3).Value
    End With
    ' Create the first section, passing in an empty string for the section name so
    ' it will always create a new set of section headers
    classRow = CreateSection(2, "", classRow)

    ' Initialize the student count and grade total for the section
    students = 0
    grades = 0
    
    ' Loop through every row in the raw data
    For rawDataRow = 2 To lastRow
        ' Check if we have a new section
        ' note: this if is why we had to initialize the section variable from the datarow
        ' we don't want to close out a section before we even begin, by initializing it to the first row
        ' we can skip directly to the Else and start processing the student data
        If section <> Worksheets("Raw Data").Cells(rawDataRow, 2).Value Then
            ' found a new section
            ' close out the previous section
            classRow = CloseSection(class, section, teacher, students, grades, summaryRow, classRow)
            ' reset the student count and grade total for the new section
            students = 0
            grades = 0
            ' increment the summary row so we know where to write the next section average
            summaryRow = summaryRow + 1
            ' save the previous class so we can check if we need to create a new sheet
            previousClass = class
            'Initialize next section with the current raw data row (from the For loop)
            With Worksheets("Raw Data")
                class = .Cells(rawDataRow, 1).Value
                section = .Cells(rawDataRow, 2).Value
                teacher = .Cells(rawDataRow, 3).Value
            End With
            ' lastly, create the new section
            classRow = CreateSection(rawDataRow, previousClass, classRow)
        Else
            'since we don't have to close out or create a new section, 
            'we can just write out the student data and keep going
            grades = grades + WriteOutStudent(rawDataRow, classRow)
            ' since WriteOutStudent returns the final grade for the student, 
            ' we have to increment the classrow and student count here
            classRow = classRow + 1
            students = students + 1
        End If
    Next rawDataRow
    ' we just exited the loop after writing the last student row
    ' so we still need to close the last section of the last class in the raw data
    classRow = CloseSection(class, section, teacher, students, grades, summaryRow, classRow)
    
    ' Now that we've processed all the data, we need to calculate the class averages across all sections
    UpdateSummarySheet summaryRow
End Sub

Sub UpdateSummarySheet(lastRow As Integer)
    ' This subroutine will calculate the class averages for each class 
    ' across all that class's sections
    ' It will insert a new row for each class average
    ' and format the row so it stands out
    ' Parameters:
    ' lastRow - the last row in the summary sheet, this will be the bottom of the loop

    ' Initialize variables
    Dim currRow As Integer
    Dim class As String
    Dim amount As Double
    Dim avg As Double
    Dim totalRows As Integer
    Dim classFirstRow As Integer
    
    ' we're going to insert rows, so we need to keep track of how many rows we've inserted
    ' so we will operate on the right row. The right row will always be the current row 
    ' plus the total number of rows that we've inserted so far
    ' so far, we haven't inserted any rows, so totalRows is 0
    totalRows = 0
    ' we'll also need to keep track of the first row of the class so we can calculate the 
    ' overall average for each class. This will be used in the Subtotal formula as the top 
    ' end of the range. The first class row is the second row in the summary sheet 
    classFirstRow = 2
    ' activate the summary sheet so we can use Range methods
    Worksheets("Summary").Activate
    ' loop through every row in the summary sheet
    For currRow = 2 To lastRow
        'use With again to reduce typing
        With Worksheets("Summary")
            ' get the class and the amount from the current row + any total rows we've inserted so far
            class = .Cells(currRow + totalRows, 1).Value
            amount = .Cells(currRow + totalRows, 4).Value
            ' look ahead to check the class in the next row
            ' if it's a different class, then we need to insert a new totals row
            If .Cells(currRow + totalRows + 1, 1).Value <> class Then
                'it's a different class so we need to insert a new row below the current row
                ' we could just add 1 to currRow + totalRows, but eliminating the +1 by
                ' putting it in the totalRows is more convenient and less error-prone
                totalRows = totalRows + 1
                ' Now use the range method for Rows to insert a new row above 
                ' the currRow + the new value for TotalRows
                Rows((currRow + totalRows) & ":" & (currRow + totalRows)).Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                ' now we need to write out the class average
                .Cells(currRow + totalRows, 1).Value = "Class Average"
                ' this next bit is the tricky bit. we need to write out the class average
                ' using the subtotal function. The first argument is 1, which means we're
                ' calculating the average. The second argument is the range of cells to average
                ' in this case, it's the range of cells in the 4th column for the class
                ' starting at the classFirstRow and ending at the current row + totalRows - 1 
                ' which is one row above the new total row we're writing to
                With .Cells(currRow + totalRows, 4)
                    .Formula = "=Round(Subtotal(1,D" & classFirstRow & ":D" & (currRow + totalRows - 1) & "),2)"
                    ' of course, we need to make the class average stand out
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
                'finally set the next classFirstRow to the row after the totals row we're editing
                classFirstRow = currRow + totalRows + 1
            ' Note: there's no Else here because if the next row is still 
            ' the same class, we don't need to do anything
            End If
        End With
        ' we're done with the current row, so we'll move on to the next one
    Next currRow
    ' OK. We've added all the class averages, now we need to autofit the columns
    Worksheets("Summary").Columns("A:D").EntireColumn.AutoFit
    ' and Activate the Summary tab so the user can see the results
    Worksheets("Summary").Activate
End Sub
