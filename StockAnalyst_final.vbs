Attribute VB_Name = "Module2"
Sub StockAnalyst()

'-- Dim variables --

Dim i As Double
Dim j As Double

Dim Ticker_Symbol As String
Dim Stock_Open As Double
Dim Stock_Close As Double
Dim Stock_Vol As Double
Dim Stock_Total As Double
Dim Stock_List As Double
Dim Year_Open As Double
Dim Year_Close As Double
Dim Year_Change As Double
Dim Per_Change As Double
Dim Summary_Row As Double

'-- Define end of stock list --

Stock_List = Range("A:A").SpecialCells(xlLastCell).Row

'-- Set up summary table --

Cells(1, 9) = "Ticker "
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percentage Change"
Cells(1, 12) = "Total Stock Volume"
Summary_Row = 2

'-- Extract ticker info --

For i = 2 To Stock_List
    
    Ticker_Symbol = Cells(i, 1)
    Stock_Open = Cells(i, 3)
    Stock_Close = Cells(i, 6)
    Stock_Vol = Cells(i, 7)

    '-- Identify end of ticker and calculate ticker info --

    If Cells(i + 1, 1) <> Cells(i, 1) Then
        
        Stock_Total = Stock_Total + Stock_Vol
        
        Year_Close = Stock_Close
        
        Year_Change = Year_Close - Year_Open
        
        If Year_Open <> 0 Then
        
            Per_Change = (Year_Close / Year_Open) - 1
            
        Else
        
            Per_Change = 0
            
        End If
        
        '-- Store ticker info in summary table --

        Cells(Summary_Row, 9) = Ticker_Symbol
        Cells(Summary_Row, 10) = Year_Change
        Cells(Summary_Row, 11) = Per_Change
        Cells(Summary_Row, 12) = Stock_Total

        Summary_Row = Summary_Row + 1
        
        Stock_Total = 0

    '-- Store ticker year open and add stock volume --

     Else
            
        If Cells(i - 1, 1) <> Cells(i, 1) Then
    
            Year_Open = Stock_Open
        
        End If

        Stock_Total = Stock_Total + Stock_Vol
        
    End If

Next i

'-- Formatting the summary table --

Range("I1:L1").HorizontalAlignment = xlCenter
Range("I1:L1").Font.Bold = True
Range("I1:L1").Columns.AutoFit
Range("J2:K" & Summary_Row).HorizontalAlignment = xlCenter
Range("J2:J" & Summary_Row).NumberFormat = "$#,##0.00"
Range("K2:K" & Summary_Row).NumberFormat = "0.00%"
Range("L2:L" & Summary_Row).NumberFormat = "#,##0"

'-- Conditional formatting  --

For j = 2 To Summary_Row

    If Cells(j, 10) > 0 Then
        
        Cells(j, 10).Interior.ColorIndex = 4
        
    ElseIf Cells(j, 10) < 0 Then
        
        Cells(j, 10).Interior.ColorIndex = 3
            
    End If

Next j

'-- Dim variables --

Dim Increase_Value As Double
Dim Decrease_Value As Double
Dim Volume_Value As Double
Dim Increase_Ticker As String
Dim Decrease_Ticker As String
Dim Volume_Ticker As String
 
 '-- Extract greatest info --

Increase_Value = 0
Decrease_Value = 0
Volume_Value = 0

 For j = 2 To Summary_Row
 
    If Cells(j, 11) > Increase_Value Then
 
        Increase_Value = Cells(j, 11)
        Increase_Ticker = Cells(j, 9)
        
    ElseIf Cells(j, 11) < Decrease_Value Then
     
        Decrease_Value = Cells(j, 11)
        Decrease_Ticker = Cells(j, 9)
        
    End If
    
    If Cells(j, 12) > Volume_Value Then
 
        Volume_Value = Cells(j, 12)
        Volume_Ticker = Cells(j, 9)
        
    End If
     
 Next j
    
'-- Set up greatest info table --

Cells(1, 15) = "Ticker"
Cells(1, 16) = "Value"
Cells(2, 14) = "Greatest % Increase"
Cells(3, 14) = "Greatest % Decrease"
Cells(4, 14) = "Greatest Total Volume"
    
'-- Populate greatest info table --
    
Cells(2, 15) = Increase_Ticker
Cells(2, 16) = Increase_Value
Cells(3, 15) = Decrease_Ticker
Cells(3, 16) = Decrease_Value
Cells(4, 15) = Volume_Ticker
Cells(4, 16) = Volume_Value

'-- Formatting greatest info table --

Range("N1:P4").Columns.AutoFit
Range("O1:P1").Font.Bold = True
Range("O1:P1").HorizontalAlignment = xlCenter
Range("O2:O4").HorizontalAlignment = xlCenter
Range("P2:P3").NumberFormat = "0.00%"
Range("P4").NumberFormat = "#,##0"

End Sub
