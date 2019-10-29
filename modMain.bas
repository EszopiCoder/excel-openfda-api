Attribute VB_Name = "modMain"
Option Explicit

Public Sub openSearchForm()

    formNDC.Show

End Sub

Public Sub resetSheet()
    
    ' Clears all formatting and restores standard height/width
    
    With ActiveSheet
        .Cells.Delete
        .Cells.ClearFormats
    End With
    ActiveWindow.FreezePanes = False
    
End Sub

Public Sub FastMode(ByVal Toggle As Boolean)
    Application.ScreenUpdating = Not Toggle
    Application.EnableEvents = Not Toggle
    Application.DisplayAlerts = Not Toggle
    Application.Calculation = IIf(Toggle, xlCalculationManual, xlCalculationAutomatic)
End Sub
