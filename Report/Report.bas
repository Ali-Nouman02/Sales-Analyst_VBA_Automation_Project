Attribute VB_Name = "Report"
Option Explicit
Dim myRegion As String, myDate As String
Dim ReplaceFile As VbMsgBoxResult
Dim Overwrite(1) As Boolean

Sub Report()
    Dim myTable As ListObject
    Dim Lastrow As Long, shall_Lastrow As Long, shregion_nextrow As Long
    
    Call Entry_Point
    Overwrite(0) = True
    Overwrite(1) = True
    
    On Error GoTo Handle
    
    Set myTable = shRegion.ListObjects("TableTemp")
    
    
    'find the last row in Shregion tab
    Lastrow = shRegion.Cells(Rows.Count, 1).End(xlUp).row
    
    If myTable.Range.Rows.Count > 2 Then
        shRegion.Activate
        myTable.DataBodyRange.Rows("2:" & Lastrow - 1).Delete
    End If
        
    myRegion = shStart.AxRegion.Value
    myDate = shAll.Cells(2, 3).Value
    myDate = Format(myDate, "YYYYMM")
    Debug.Print myDate
    Debug.Print myRegion
    
    
    'using formulas to find find which region each row belongs to
    shall_Lastrow = shAll.Cells(shAll.Rows.Count, 1).End(xlUp).row
    
    'find the respective region and pull down the formula to the last row
    shAll.Range("I2").FormulaR1C1 = "=VLOOKUP(RC[-8],MCompany,4,FALSE)"
    shAll.Range("I2").AutoFill Destination:=shAll.Range("I2:I" & shall_Lastrow)
    shAll.Application.Calculate
    
    
    'filter the row
    shAll.Range("A1:I" & shall_Lastrow).AutoFilter Field:=9, Criteria1:=myRegion
    
    'copy the visible rows
    shAll.Range("A2:I" & shall_Lastrow).SpecialCells(xlCellTypeVisible).Copy
    
    'shregion_nextrow = shRegion.Cells(shRegion.Rows.Count, 1).End(xlUp).Row + 1
    shRegion.Range("A2").PasteSpecial xlPasteValues
    
    'new lastrow in the shregion tab
    Lastrow = shRegion.Cells(Rows.Count, 1).End(xlUp).row
    
    'correct the formating of the table in the shregion tab
    shRegion.Range("I3:I31").Copy
    shRegion.Range("A2:H" & Lastrow).PasteSpecial Paste:=xlPasteFormats
    shRegion.Application.CutCopyMode = False
    
    'Remove the filter on shall tab and clear the values in the column I
    shAll.Columns("I:I").AutoFilter
    shAll.Columns("I:I").ClearContents
    
    'Update the value in U1 Cell
    shRegion.Range("U1").Value = myDate
    shRegion.PivotTables("PTRegion").PivotCache.Refresh
    
    Call Region_Report
    Call Manager_Report
    
    'should we overwrite the files
    If Overwrite(0) = False And Overwrite(1) = False Then
        MsgBox "No Files were created", , "Operation Cancelled"
    Else
        MsgBox "File(s) created in the same directory as this workbook", , "Welldone"
    End If
    
    
    Call Exit_Point
    
    Exit Sub
Handle:
    If Err.Number = 1004 Then
        MsgBox "It seems that data for this region has not been uploaded yet"
        shAll.Range("A1").AutoFilter
        shAll.Columns("i").Delete
        Exit Sub
'    Else
'        MsgBox "Looks like no Data has been uploaded yet"
'        Exit Sub
    End If
    Call Exit_Point
End Sub

Private Sub Region_Report()
    Dim NewRBook As Workbook
    Dim newRPath As String
    Dim PT As PivotTable
    
    newRPath = ThisWorkbook.Path & "\" & "RM_" & myDate & myRegion & ".xlsx"
    
    If Len(Dir(newRPath)) > 0 Then 'if true file exists
        ReplaceFile = MsgBox("Regional File Exists. Do you want to overwrite it?", vbOKCancel, "Overwrite?")
        If ReplaceFile = vbCancel Then
            Overwrite(0) = False
            Exit Sub
        End If
    End If
    
    
    Set NewRBook = Workbooks.Add
    
    shRegion.Copy Before:=NewRBook.Sheets(1)
    
    'remove the common filter for slicer
    NewRBook.SlicerCaches("Slicer_Company_Name").PivotTables.RemovePivotTable _
    (ActiveSheet.PivotTables("PTArticle"))
    
    'change pivot cache
    NewRBook.ActiveSheet.PivotTables("PTRegion").ChangePivotCache _
    NewRBook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="TableTemp")
    
    For Each PT In NewRBook.ActiveSheet.PivotTables
        If PT.Name <> "PTRegion" Then
            PT.ChangePivotCache ("PTRegion")
        End If
    Next PT

    'add back the common slicer
    NewRBook.SlicerCaches("Slicer_Company_Name").PivotTables.AddPivotTable _
    (ActiveSheet.PivotTables("PTArticle"))
    
    
    'clean up
    With ActiveSheet
        .Columns("AJ:AZ").Delete
        .Range("TableTemp").CurrentRegion.Columns.Group
        .Range("TableTemp").EntireColumn.Hidden = True
        .Range("A1").CurrentRegion.Copy
        .Range("A1").PasteSpecial (xlPasteValues)
        
    End With
    
    NewRBook.SaveAs Filename:=newRPath
    NewRBook.Close
    
    TaskRow = shStart.Range("B" & shStart.Rows.Count).End(xlUp).row + 1
    shStart.Range("B" & TaskRow).Value = "Regional Report for" & myRegion & myDate & " - created on:" & Now
      
End Sub
Private Sub Manager_Report()
    Dim newBook As Workbook
    Dim newPath As String
    
    newPath = ThisWorkbook.Path & "\" & "M_" & myDate & myRegion & ".xlsx"
    
    If Len(Dir(newPath)) > 0 Then 'if true file exists
        ReplaceFile = MsgBox("Manager File Exists. Do you want to overwrite it?", vbOKCancel, "Overwrite?")
        If ReplaceFile = vbCancel Then
            Overwrite(1) = False
            Exit Sub
        End If
    End If
    
    Set newBook = Workbooks.Add
    ActiveWindow.DisplayGridlines = False
    
    shRegion.PivotTables("PTManager").TableRange1.Copy
    
    With newBook.Sheets(1).Range("A1")
        .PasteSpecial (xlPasteValues)
        .PasteSpecial (xlPasteFormats)
        .PasteSpecial (xlPasteColumnWidths)
    End With
    
    'formating to make it look better
    newBook.Sheets(1).Range("1:2").EntireRow.Insert
    With newBook.Sheets(1).Range("A1")
        .Value = "Regional Manager"
        .Font.Color = vbWhite
        .RowHeight = 27.75
        .Font.Bold = True
    End With
    
    
    With newBook.Sheets(1).Range("A1:J1").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    With newBook.Sheets(1).Range("A1:J1")
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    With newBook.Sheets(1).Range("A1:J1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    
    newBook.SaveAs Filename:=newPath
    newBook.Close
    
    'document tasks completed
    TaskRow = shStart.Range("B" & shStart.Rows.Count).End(xlUp).row + 1
    shStart.Range("B" & TaskRow).Value = "Manager Report for" & myRegion & myDate & " - created on:" & Now
    
    
End Sub

Sub Create_CSV()
Dim rng As Range
    Dim cell As Range, row As Range
    Dim outputFile As String, rowData As String
    Dim fileNumber As Integer
    Dim Lastrow As Long
    
    Call Entry_Point
    
    Lastrow = shAll.Cells(Rows.Count, 1).End(xlUp).row
    
    ' Define the range you want to export
    Set rng = shAll.Range("A2:H" & Lastrow)

    ' Specify the output file path and name
    outputFile = ThisWorkbook.Path & "\" & Format(shAll.Range("C2").Value, "YYYYMM") & ".csv"
    
    fileNumber = FreeFile

    ' Open the file for writing
    Open outputFile For Output As #fileNumber
    
    'first loop through each row in the rng
    For Each row In rng.Rows
        rowData = "" ' important :Reset rowData for each row
        
        'then loop through each cell in the row
        For Each cell In row.Cells
            rowData = rowData & cell.Value & ";"
        Next cell

        ' Remove the last semicolon
        rowData = Left(rowData, Len(rowData) - 1)
        
        'write the data to the csv file
        Print #fileNumber, rowData
    Next row

    ' Close the file
    Close #fileNumber

    ' Notify the user
    MsgBox "CSV file has been saved in the same directory as this workbook", vbInformation, "CSV Saved"
    
    'document the tasks completed
    TaskRow = shStart.Range("B" & shStart.Rows.Count).End(xlUp).row + 1
    shStart.Range("B" & TaskRow).Value = "CSV for month" & myDate & " - created on:" & Now
    

    Call Exit_Point
    
End Sub
