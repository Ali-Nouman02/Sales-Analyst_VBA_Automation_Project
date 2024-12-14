Attribute VB_Name = "Ímport"
Option Explicit

Sub Import_Sales()

    Dim filePath As Variant
    Dim cnt As Byte, c As Byte
    Dim fileToOpen As Workbook
    Dim nextRow As Long, lastTempRow As Long
    Const startTempRow As Byte = 4
    Dim FindCell As Range
    Dim FindValue As String
    Dim CompanyImported() As String
    
    Call Entry_Point
    
    On Error GoTo Handle
    filePath = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx ", _
                                                        Title:="Please Select the required Files", _
                                                        MultiSelect:=True)
    'check if the user selected any files
    If IsArray(filePath) Then
        For cnt = 1 To UBound(filePath)
            'clear the data on newData Sheet
            shNewDat.Cells.Clear
            
            'opens the selected workbook
            Set fileToOpen = Workbooks.Open(Filename:=filePath(cnt))
            
            'Copy and the paste the values
            fileToOpen.Worksheets("Sales").Cells.Copy
            shNewDat.Range("A1").PasteSpecial xlPasteValues
            
            fileToOpen.Application.CutCopyMode = False
            
            'close the file
            fileToOpen.Close
            
            
            'copy the company code to the summary tab
            lastTempRow = shNewDat.Cells(Rows.Count, 2).End(xlUp).row
            nextRow = shAll.Cells(Rows.Count, 1).End(xlUp).row + 1
            
            shAll.Range("A" & nextRow & ":A" & ((nextRow + (lastTempRow - startTempRow)) - 1)) = shNewDat.Cells(2, 3).Value
            
            'fill the other relevant columns
            c = 2
            Do While shAll.Cells(1, c).Value <> ""
                Debug.Print shAll.Cells(1, c).Value
                FindValue = shAll.Cells(1, c).Value
                
                Set FindCell = shNewDat.Rows(startTempRow).Find(What:=FindValue, _
                                                                LookIn:=xlValues, _
                                                                LookAt:=xlWhole)
                Debug.Print FindCell.Column
                If Not FindCell Is Nothing Then
                    shNewDat.Range(shNewDat.Cells(startTempRow + 1, FindCell.Column), shNewDat.Cells(lastTempRow, FindCell.Column)).Copy
                    
                    shAll.Cells(nextRow, c).PasteSpecial Paste:=xlPasteValues
                    
                End If
                c = c + 1
            Loop
              'give a size to the Array
              ReDim Preserve CompanyImported(1 To cnt)
              CompanyImported(cnt) = shNewDat.Range("C2").Value
                              
                              
        Next cnt
        MsgBox "Data Imported Sucessfully"
    End If
    shStart.Select
    Range("A1").Select
    
    'to find the row to document are completed tasks
    TaskRow = shStart.Range("B" & shStart.Rows.Count).End(xlUp).row + 1
    shStart.Range("B" & TaskRow).Value = "Data imported for:" & Join(CompanyImported, ",")
    
    
    
    Call Exit_Point
    Exit Sub
    
Handle:
    If Err.Number = 9 Then
        MsgBox "It looks like you have selectecd the wrong file"
    Else
        MsgBox "An error has occurred"
    End If
Call Exit_Point
    
End Sub



