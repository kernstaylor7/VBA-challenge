VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Stats():

Dim ticker As String
Dim total_tickers As Integer
Dim Last_row_state As Long
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim gpi As Double
Dim gpi_ticker As String
Dim gpd As Double
Dim gpd_ticker As String
Dim gsv As Double
Dim gsv_ticker As String

For Each ws In Worksheets
    ws.Activate

Last_row_state = ws.Cells(Rows.Count, "A").End(xlUp).Row
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    total_tickers = 0
    ticker = ""
    yearly_change = 0
    opening_price = 0
    percent_change = 0
    total_stock_volume = 0
    
    For i = 2 To Last_row_state

        ticker = Cells(i, 1).Value
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
    total_stock_volume = total_stock_volume + Cells(i, 7).Value
    If Cells(i + 1, 1).Value <> ticker Then
 
    total_tickers = total_tickers + 1
    Cells(total_tickers + 1, 9) = ticker
    closing_price = Cells(i, 6)
    yearly_change = closing_price - opening_price
            
    Cells(total_tickers + 1, 10).Value = yearly_change
            
    If yearly_change > 0 Then
        Cells(total_tickers + 1, 10).Interior.ColorIndex = 4
    ElseIf yearly_change < 0 Then
        Cells(total_tickers + 1, 10).Interior.ColorIndex = 3
    Else
        Cells(total_tickers + 1, 10).Interior.ColorIndex = 6
    End If
    If opening_price = 0 Then
        percent_change = 0
    Else
        percent_change = (yearly_change / opening_price)
    End If
        opening_price = 0
            
    Cells(total_tickers + 1, 12).Value = total_stock_volume
            
    total_stock_volume = 0
    End If
        
    Next i
    
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

Last_row_state = ws.Cells(Rows.Count, "I").End(xlUp).Row

gpi = Cells(2, 11).Value
gpi_ticker = Cells(2, 9).Value
gpd = Cells(2, 11).Value
gpd_ticker = Cells(2, 9).Value
gsv = Cells(2, 12).Value
gsv_ticker = Cells(2, 9).Value

For i = 2 To Last_row_state
    If Cells(i, 11).Value > gpi Then
        gpi = Cells(i, 11).Value
        gpi_ticker = Cells(i, 9).Value
    End If
        If Cells(i, 11).Value < gpd Then
            gpd = Cells(i, 11).Value
            gpd_ticker = Cells(i, 9).Value
    End If
        If Cells(i, 12).Value > gsv Then
            gsv = Cells(i, 12).Value
            gsv_ticker = Cells(i, 9).Value
    End If
        
    Next i
   
    Range("P2").Value = Format(gpi_ticker, "Percent")
    Range("Q2").Value = Format(gpi, "Percent")
    Range("P3").Value = Format(gpd_ticker, "Percent")
    Range("Q3").Value = Format(gpd, "Percent")
    Range("P4").Value = gsv_ticker
    Range("Q4").Value = gsv
    
Next ws


End Sub

