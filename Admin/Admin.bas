Attribute VB_Name = "Admin"
Option Explicit
Public TaskRow As Long
Sub Entry_Point()
    With Application
        .StatusBar = "Your Assistant is Busy Working"
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With

End Sub

Sub Exit_Point()
    With Application
        .StatusBar = ""
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .CutCopyMode = False
    End With
End Sub

Sub Clear_OldMonth()

'clear data
shAll.Range("A2", "A" & shAll.Rows.Count - 1).EntireRow.Delete
shStart.Range("B20:B33").ClearContents

End Sub
