Attribute VB_Name = "Module2"
Sub Stocktest2():

'Set Variables
Dim Ticker As String
Dim openvalue As Double
Dim closevalue As Double
Dim Total_stock_volume As Double
Dim Summary_table As Integer
Dim ticker_date As Long
Dim annual_change As Double
Dim percent_change As Double
Total_stock_volume = 0
openvalue = Cells(2, 3).Value
closevalue = 0
Summary_table = 2

'Determine rows in sheet
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'set worksheet count
ws_count = ActiveWorkbook.Worksheets.Count

Dim ws As Worksheet
'For x = 1 To ws_count
'MsgBox (ActiveWorkbook.Worksheets(x).Name)
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
Summary_table = 2

'Set up Summary table with headers
Range("J1").Value = "Ticker Symbol"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"
Range("J1:M1").Font.Bold = True

'Run through first row of data to get tickers
'For Each ws In ThisWorkbook.Worksheets
    'ws.Activate
    
    For i = 2 To lastrow
    If Cells(i + 1, 1).Value <> Cells(i, 1) Then
    
    'set the ticker symbol
    Ticker = Cells(i, 1).Value
    'MsgBox (Cells(i, 1).Value)
    
    'set the total stock volume
    Total_stock_volume = Total_stock_volume + Cells(i, 7).Value
    
    'set closing value
    closevalue = Cells(i, 6).Value
    
    'set the yearly change value
    Range("K" & Summary_table).Value = closevalue - openvalue
    
    If openvalue = 0 Then
    percent_change = 0
    Else
    percent_change = (closevalue - openvalue) / openvalue
    End If
    
    'set open value
    openvalue = Cells(i + 1, 3).Value
    
    'Print the volume total in summary table
    Range("L" & Summary_table).Value = percent_change
    
    'move this ticker value into summary table
    Range("J" & Summary_table).Value = Ticker
    
    'Print the volume total in summary table
    Range("M" & Summary_table).Value = Total_stock_volume
    
    'Add row
    Summary_table = Summary_table + 1
    
    'reset volume
    Total_stock_volume = 0
    
    Else
    Total_stock_volume = Total_stock_volume + Cells(i, 7).Value
    If openvalue = 0 Then
    openvalue = Cells(i, 3).Value
    End If
    End If
    
Next i


For i = 2 To lastrow

' set percent change to percent format
    Cells(i, 12).NumberFormat = "0.00%"
    
' set color code for red for items in negative in yearly change
    If Cells(i, 11).Value < 0 Then
    Cells(i, 11).Interior.ColorIndex = 3
    
'set color code for green for items in positive in yearly change
    Else
    Cells(i, 11).Interior.ColorIndex = 4
    End If

    
Next i
Next ws



End Sub
