VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8724.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ImportButton_Click()
Dim DateDay As String
Dim DateMonth As String
Dim DateYear As String
Dim firstslash As Integer
Dim url As String
Dim Currency1, Currency2 As String

Sheets("Sheet3").Visible = True
Sheets("Sheet1").Visible = True

Call ApplicationSpeedOptimize

DateDay = Format(DateBox, "dd")
DateMonth = Format(DateBox, "mm")
DateYear = Format(DateBox, "yyyy")


Currency1 = Left(ItemBox, 3)
Currency2 = Left(ItemBox2, 3)

    url = "URL;https://www.xe.com/currencytables/?from=" & Currency1 & "&date=" & DateYear & "-" & DateMonth & "-" & DateDay
    With Worksheets("Sheet1").QueryTables.Add(Connection:=url, Destination:=Worksheets("Sheet1").Range("A1"))
        .Name = "My Query"
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With


For i = 0 To 167
    If Sheets("sheet1").Cells(i + 15, 1) = Currency2 Then
       TextConverted = TextAmount * Sheets("sheet1").Cells(i + 15, 3)
    End If
Next
Sheets("sheet1").Cells.Clear
Sheets("Sheet3").Visible = False
Sheets("Sheet1").Visible = False

Call ApplicationRestoreAfterSpeedOptimize
End Sub


Private Sub Label2_Click()

End Sub

Private Sub QuitBox_Click()
Unload Me

End Sub
''''problem other computer with US setting will print the date in the wrong order
Private Sub ToggleButton1_Click()
Dim TodayDate As String
Dim i As Integer

Dim DateDay As String
Dim DateMonth As String
Dim DateYear As String
Dim firstslash As Integer
Dim url As String
Dim DateArray As Variant
Dim Currency1, Currency2 As String

Call ApplicationSpeedOptimize

Sheets("Sheet3").Visible = True
Sheets("Sheet1").Visible = True
Sheets("Sheet3").Select
Range("a30").Select
TodayDate = DateBox
For i = 30 To 1 Step -1
    Range("A" & 30 - i + 1) = DateAdd("d", -i + 1, TodayDate)
Next i


For i = 30 To 1 Step -1

DateDay = Format(Cells(30 - i + 1, 1), "dd")
DateMonth = Format(Cells(30 - i + 1, 1), "mm")
DateYear = Format(Cells(30 - i + 1, 1), "yyyy")

Currency1 = Left(ItemBox, 3)
Currency2 = Left(ItemBox2, 3)

    url = "URL;https://www.xe.com/currencytables/?from=" & Currency1 & "&date=" & DateYear & "-" & DateMonth & "-" & DateDay
    With Worksheets("Sheet1").QueryTables.Add(Connection:=url, Destination:=Worksheets("Sheet1").Range("A1"))
        .Name = "My Query"
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    
Sheets("sheet1").Select
        
        For j = 0 To 167
            If Sheets("sheet1").Cells(j + 15, 1) = Currency2 Then
               Sheets("sheet3").Cells(30 - i + 1, 2) = TextAmount * Sheets("sheet1").Cells(j + 15, 3)
            End If
        Next j
Sheets("sheet1").Cells.Clear
Sheets("sheet3").Select
Cells(1, 1).Select

Next i
Call ApplicationRestoreAfterSpeedOptimize

    Sheets("sheet3").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet3!$A$1:$B$30")
    ActiveChart.Location Where:=xlLocationAsNewSheet
    ActiveChart.ChartTitle.Text = "Last 30 days"
    
Sheets("Sheet3").Visible = False
Sheets("Sheet1").Visible = False
 Unload Me
End Sub

