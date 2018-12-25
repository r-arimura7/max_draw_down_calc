Attribute VB_Name = "Module4"
Option Explicit

Type DrawDownResult
    startDateRow As Long
    endDateRow As Long
    value As Double
End Type

Const ORIGINAL_DATA_START_ROW As Long = 3
Const FUND_START_COLUMN As Long = 2
Const MONTH_COLUMN As Long = 1
Const FUND_DATA_START_ROW = 1
Const FUND_DATA_START_COL = 2
Const TOP_LEFT_CELL_ROW = 1
Const TOP_LEFT_CELL_COLUMN = 1
Const PREPROCESSED_MONTH_START_ROW = 2
Const RESULT_1Y_VALUE_ROW = 3
Const RESULT_1Y_MAXDD_START_ROW = 4
Const RESULT_1Y_MAXDD_END_ROW = 5
Const RESULT_3Y_VALUE_ROW = 8
Const RESULT_3Y_MAXDD_START_ROW = 9
Const RESULT_3Y_MAXDD_END_ROW = 10
Const RESULT_5Y_VALUE_ROW = 13
Const RESULT_5Y_MAXDD_START_ROW = 14
Const RESULT_5Y_MAXDD_END_ROW = 15
Const RESULT_ITD_VALUE_ROW = 18
Const RESULT_ITD_MAXDD_START_ROW = 19
Const RESULT_ITD_MAXDD_END_ROW = 20

Sub main()
  Maxdd_output
End Sub

Sub Maxdd_output()
  Dim originalSheet As Worksheet
  Dim preprocessedSheet As Worksheet
  Dim resultSheet As Worksheet
  Dim durationMonth As Variant
  Dim fundcolNum As Long
  Dim maxddResult As DrawDownResult
    
  Set originalSheet = Worksheets("Original")
  Set preprocessedSheet = Worksheets("Data")
  Set resultSheet = Worksheets("MDD")
  
  Dim DURATION_MONTHS() As Variant
  DURATION_MONTHS = Array(12, 36, 60, 0) ' 0 means "ITD"
  
  preprocess originalSheet, preprocessedSheet, resultSheet
  For Each durationMonth In DURATION_MONTHS
    For fundcolNum = FUND_DATA_START_COL To getFundColNums(preprocessedSheet)
      Debug.Print "a"
      maxddResult = calcMaxDD(preprocessedSheet, fundcolNum, durationMonth)
      printMaxDD resultSheet, durationMonth, maxddResult, fundcolNum
      resultSheet.Select
    Next
  Next
End Sub

Sub preprocess(originalSheet As Worksheet, preprocessedSheet As Worksheet, resultSheet As Worksheet)
  invertAndCopyMonthColumn originalSheet, preprocessedSheet
  copyFundNames originalSheet, preprocessedSheet, resultSheet
  invertAndCopyPerf originalSheet, preprocessedSheet
  formatvalues preprocessedSheet
End Sub

Sub invertAndCopyMonthColumn(originalSheet As Worksheet, preprocessedSheet As Worksheet)
  Dim lastRow As Long
  Dim currentRowNum As Long
  Dim preprocessedRowNum As Long
    lastRow = originalSheet.Cells(ORIGINAL_DATA_START_ROW, MONTH_COLUMN).End(xlDown).Row
    
    For currentRowNum = ORIGINAL_DATA_START_ROW To lastRow
          preprocessedRowNum = lastRow + ORIGINAL_DATA_START_ROW - 1 - currentRowNum
          preprocessedSheet.Cells(preprocessedRowNum, MONTH_COLUMN).value = originalSheet.Cells(currentRowNum, MONTH_COLUMN).value
    Next
End Sub

Sub copyFundNames(originalSheet As Worksheet, preprocessedSheet As Worksheet, resultSheet As Worksheet)
    Dim lastFundCol As Long
    Dim fundNum As Long
    Dim currentColumn As Long
    
    lastFundCol = originalSheet.Cells(FUND_DATA_START_ROW, FUND_DATA_START_COL).End(xlToRight).Column
      
    For currentColumn = FUND_DATA_START_COL To lastFundCol
      preprocessedSheet.Cells(FUND_DATA_START_ROW, currentColumn).value = originalSheet.Cells(FUND_DATA_START_ROW, currentColumn).value
      resultSheet.Cells(FUND_DATA_START_ROW, currentColumn).value = originalSheet.Cells(FUND_DATA_START_ROW, currentColumn).value
    Next
End Sub

Sub invertAndCopyPerf(originalSheet As Worksheet, preprocessedSheet As Worksheet)
  Dim lastRow As Long
  Dim lastCol As Long
  Dim currentCol As Long
  Dim currentRow As Long
  Dim preprocessedRow As Long
  
  lastRow = originalSheet.Cells(ORIGINAL_DATA_START_ROW, MONTH_COLUMN).End(xlDown).Row
  lastCol = originalSheet.Cells(FUND_DATA_START_ROW, FUND_DATA_START_COL).End(xlToRight).Column
  
  For currentCol = FUND_DATA_START_COL To lastCol
    For currentRow = ORIGINAL_DATA_START_ROW To lastRow
      preprocessedRow = lastRow + ORIGINAL_DATA_START_ROW - 1 - currentRow
      preprocessedSheet.Cells(preprocessedRow, currentCol).value = originalSheet.Cells(currentRow, currentCol).value
    Next
  Next
End Sub

Sub formatvalues(preprocessedSheet As Worksheet)
Dim lastRow As Long
Dim lastCol As Long

  lastRow = preprocessedSheet.Cells(PREPROCESSED_MONTH_START_ROW, MONTH_COLUMN).End(xlDown).Row
  lastCol = preprocessedSheet.Cells(FUND_DATA_START_ROW, FUND_DATA_START_COL).End(xlToRight).Column
    
  preprocessedSheet.Select
  preprocessedSheet.Range(Cells(TOP_LEFT_CELL_ROW, TOP_LEFT_CELL_COLUMN), Cells(lastRow, lastCol)).Select
        
        With Selection.Font
        .Name = "Calibri"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Cells(FUND_DATA_START_ROW, FUND_DATA_START_COL).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormatLocal = "[$-409]mmmm yy;@"
    Selection.NumberFormatLocal = "[$-409]mmm yy;@"
    Cells(FUND_DATA_START_ROW, FUND_DATA_START_COL).EntireColumn.ColumnWidth = 9.75
    
    Cells(FUND_DATA_START_ROW, FUND_DATA_START_COL).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Percent"
    Selection.NumberFormatLocal = "0.00%"

End Sub

Function getFundColNums(preprocessedSheet As Worksheet) As Long
  Dim lastCol As Long
  
  lastCol = preprocessedSheet.Cells(FUND_DATA_START_ROW, FUND_DATA_START_COL).End(xlToRight).Column
  getFundColNums = lastCol
  
End Function

Function calcMaxDD(sHeet As Worksheet, fundcolNum As Long, durationMonth As Variant) As DrawDownResult

    Dim currentResult As DrawDownResult
    Dim worstResult As DrawDownResult
    Dim currentRate As Double
    Dim rate As Double
    Dim currentRow As Long
    Dim startRow As Long
    Dim lastRow As Long
       
    lastRow = sHeet.Cells(PREPROCESSED_MONTH_START_ROW, MONTH_COLUMN).End(xlDown).Row
    startRow = getStartRowByColumn(sHeet, fundcolNum, lastRow, durationMonth)
    
    currentRate = 1
    
    worstResult.value = 1
    worstResult.startDateRow = 0
    worstResult.endDateRow = 0

    currentResult.value = 1
    currentResult.startDateRow = startRow - 1
    currentResult.endDateRow = 0

    For currentRow = startRow To lastRow
        rate = getRateAtCell(sHeet, currentRow, fundcolNum)
        currentRate = currentRate * rate
                
        If currentRate >= 1 Then
            currentRate = 1
            currentResult.startDateRow = currentRow

        ElseIf currentRate < currentResult.value Then
            currentResult.value = currentRate
            currentResult.endDateRow = currentRow
                    
            'currentRate‚ªXV‚³‚ê‚é‚Æ‚«‚ÍA‘SŒn—ñ‚Å‚ÌÅˆ«‚à“¯Žž‚É•]‰¿‚·‚é
            If currentRate < worstResult.value Then
                worstResult.value = currentRate
                worstResult.startDateRow = sHeet.Cells(currentResult.startDateRow + 1, MONTH_COLUMN)
                worstResult.endDateRow = sHeet.Cells(currentResult.endDateRow, MONTH_COLUMN)
            End If
        End If
            
    Next
    calcMaxDD = worstResult

End Function

Sub printMaxDD(resultSheet As Worksheet, durationMonth As Variant, maxddResult As DrawDownResult, fundcolNum As Long)
 
  If durationMonth = 12 Then
    If maxddResult.value >= 1 Then
      resultSheet.Cells(RESULT_1Y_VALUE_ROW, fundcolNum).value = "n.a."
      resultSheet.Cells(RESULT_1Y_MAXDD_START_ROW, fundcolNum).value = "n.a."
      resultSheet.Cells(RESULT_1Y_MAXDD_END_ROW, fundcolNum).value = "n.a."
    Else
      resultSheet.Cells(RESULT_1Y_VALUE_ROW, fundcolNum).value = maxddResult.value - 1
      resultSheet.Cells(RESULT_1Y_MAXDD_START_ROW, fundcolNum).value = maxddResult.startDateRow
      resultSheet.Cells(RESULT_1Y_MAXDD_END_ROW, fundcolNum).value = maxddResult.endDateRow
    End If
   
  ElseIf durationMonth = 36 Then
    If maxddResult.value >= 1 Then
      resultSheet.Cells(RESULT_3Y_VALUE_ROW, fundcolNum).value = "n.a."
      resultSheet.Cells(RESULT_3Y_MAXDD_START_ROW, fundcolNum).value = "n.a."
      resultSheet.Cells(RESULT_3Y_MAXDD_END_ROW, fundcolNum).value = "n.a."
    Else
      resultSheet.Cells(RESULT_3Y_VALUE_ROW, fundcolNum).value = maxddResult.value - 1
      resultSheet.Cells(RESULT_3Y_MAXDD_START_ROW, fundcolNum).value = maxddResult.startDateRow
      resultSheet.Cells(RESULT_3Y_MAXDD_END_ROW, fundcolNum).value = maxddResult.endDateRow
    End If
  
  ElseIf durationMonth = 60 Then
    If maxddResult.value >= 1 Then
      resultSheet.Cells(RESULT_5Y_VALUE_ROW, fundcolNum).value = "n.a."
      resultSheet.Cells(RESULT_5Y_MAXDD_START_ROW, fundcolNum).value = "n.a."
      resultSheet.Cells(RESULT_5Y_MAXDD_END_ROW, fundcolNum).value = "n.a."
    Else
      resultSheet.Cells(RESULT_5Y_VALUE_ROW, fundcolNum).value = maxddResult.value - 1
      resultSheet.Cells(RESULT_5Y_MAXDD_START_ROW, fundcolNum).value = maxddResult.startDateRow
      resultSheet.Cells(RESULT_5Y_MAXDD_END_ROW, fundcolNum).value = maxddResult.endDateRow
    End If
  
  ElseIf durationMonth = 0 Then
    If maxddResult.value >= 1 Then
      resultSheet.Cells(RESULT_ITD_VALUE_ROW, fundcolNum).value = "n.a."
      resultSheet.Cells(RESULT_ITD_MAXDD_START_ROW, fundcolNum).value = "n.a."
      resultSheet.Cells(RESULT_ITD_MAXDD_END_ROW, fundcolNum).value = "n.a."
    Else
      resultSheet.Cells(RESULT_ITD_VALUE_ROW, fundcolNum).value = maxddResult.value - 1
      resultSheet.Cells(RESULT_ITD_MAXDD_START_ROW, fundcolNum).value = maxddResult.startDateRow
      resultSheet.Cells(RESULT_ITD_MAXDD_END_ROW, fundcolNum).value = maxddResult.endDateRow
    End If
  End If
  
End Sub

Function getRateAtCell(sHeet As Worksheet, rowNum As Long, fundcolNum As Long) As Double
    Dim performance As Double
    performance = sHeet.Cells(rowNum, fundcolNum).value
    getRateAtCell = 1 + performance
End Function

Function getStartRowByColumn(sHeet As Worksheet, fundcolNum As Long, lastRow As Long, durationMonth As Variant) As Long
Dim startRow As Long
Dim dataStartRow As Long

    If durationMonth <> 0 Then
      startRow = lastRow - durationMonth + 1
    ElseIf durationMonth = 0 Then
      startRow = sHeet.Cells(PREPROCESSED_MONTH_START_ROW, fundcolNum).End(xlDown).Row
      
      If startRow = lastRow Then
        dataStartRow = Cells(PREPROCESSED_MONTH_START_ROW, MONTH_COLUMN).Row
        startRow = dataStartRow
      End If
    
    End If

getStartRowByColumn = startRow

End Function
