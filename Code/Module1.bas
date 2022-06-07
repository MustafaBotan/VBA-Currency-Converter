Attribute VB_Name = "Module2"

Sub opernForm()
Dim i As Integer
Dim DateArray As Variant
Dim TodayDate As String

Sheets("Sheet2").Select
Range("a1").Select
For i = 1 To WorksheetFunction.CountA(Columns("A:A"))
    UserForm1.ItemBox.AddItem ActiveCell.Offset(i - 1, 0) & " - " & ActiveCell.Offset(i - 1, 1)
    UserForm1.ItemBox2.AddItem ActiveCell.Offset(i - 1, 0) & " - " & ActiveCell.Offset(i - 1, 1)
Next i
DateArray = Split(Now())
UserForm1.DateBox = DateArray(0)
UserForm1.ItemBox.Text = Range("a1") & " - " & Range("b1")
UserForm1.ItemBox2.Text = Range("a2") & " - " & Range("b2")
UserForm1.TextAmount.Value = 1
UserForm1.Show

End Sub

Sub ApplicationSpeedOptimize()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual   ' turn off the automatic calculation
    Application.DisplayStatusBar = False            ' turn off status bar updates
    Application.EnableEvents = False                ' ignore events
    ActiveSheet.DisplayPageBreaks = False           ' don't use time to calculate page breaks
End Sub

Sub ApplicationRestoreAfterSpeedOptimize()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
End Sub
