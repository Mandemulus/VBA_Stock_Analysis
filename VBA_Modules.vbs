
Private Sub For_Loop()

Dim WS_Count As Integer     ' how many worksheets
Dim NumRows As String       ' how many rows in worksheet
Dim I As Integer            ' loop thru worksheets
Dim II As Variant           ' loop thru rows in worksheet
Dim III As Integer          ' write to worksheet

Dim stock_Ticker As String  ' contains the stock ticker we are currently processing
Dim ticker_Date  As String  ' contains the date the
Dim open_Price   As Double ' contains the stocks openning price
Dim high_Price   As Double  ' contains the stocks daily high price
Dim low_Price    As Double  ' contains the stocks daily low price
Dim close_Price  As Double ' contains the stocks closing price
Dim volume       As Double ' contains the stocks days volume
   
    ' Set WS_Count equal to the number of worksheets in the active workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count
    ' Begin the loop. Loop thru Worksheets
    For I = 1 To WS_Count
        ' Sort Worksheet to ensure we get all the data
        With Worksheets(ActiveWorkbook.Worksheets(I).Name)
            With .Cells(1, "A").CurrentRegion
                .Cells.Sort Key1:=.Range("A1"), Order1:=xlAscending, _
                            Key2:=.Range("B1"), Order2:=xlAscending, _
                            Orientation:=xlTopToBottom, Header:=xlYes
            End With    ' sort the Active Worksheet
        End With        ' with the Active Worksheet

        III = 1     ' This is used when writing to the worksheet
        ' clean up from previous runs if any
        Worksheets(ActiveWorkbook.Worksheets(I).Name).Columns(9).ClearContents
        Worksheets(ActiveWorkbook.Worksheets(I).Name).Columns(10).ClearContents
        Worksheets(ActiveWorkbook.Worksheets(I).Name).Columns(11).ClearContents
        Worksheets(ActiveWorkbook.Worksheets(I).Name).Columns(12).ClearContents
        ' populate the headers
        Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("I1").Value = "Ticker"
        Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("J1").Value = "Yearly Change"
        Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("K1").Value = "Percentage Change"
        Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("L1").Value = "Total Stock Volume"
 

      ' Set numrows = number of rows in Active Worksheet
        NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
        For II = 2 To NumRows ' start with 2 so we skip the header line
            If II = 2 Then     ' we are at the first row so move all columns to variables
                stock_Ticker = Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("A" & II).Value
                ticker_Date = Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("B" & II).Value
                open_Price = Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("C" & II).Value
                high_Price = Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("D" & II).Value
                low_Price = Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("E" & II).Value
                close_Price = Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("F" & II).Value
                volume = Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("G" & II).Value
            
            ' If we have come to a different ticker symbol write line to worksheet and reload variables
            ElseIf Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("A" & II).Value <> stock_Ticker Then
                III = III + 1
                ' we have come to the beginning of a new ticker so write out the current symbol
                Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("I" & III).Value = stock_Ticker
                
                ' yearly change
                Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("J" & III).Value = close_Price - open_Price
                If (close_Price - open_Price) >= 0 Then
                    Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("J" & III).Interior.ColorIndex = 4
                Else
                    Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("J" & III).Interior.ColorIndex = 3
                End If
                
                ' percentage change Round(((close_Price - open_Price) / open_Price) * 100, 2)
                If open_Price = 0 Then ' cannot divide by 0
                   Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("K" & III).Value = 0 & "%"
                Else
                    Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("K" & III).Value = _
                        Round(((close_Price - open_Price) / open_Price) * 100, 2) & "%"
                End If
                
                Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("L" & III).Value = volume
            
                ' move newest values to variables
                stock_Ticker = Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("A" & II).Value
                ticker_Date = Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("B" & II).Value
                open_Price = Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("C" & II).Value
                high_Price = Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("D" & II).Value
                low_Price = Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("E" & II).Value
                close_Price = Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("F" & II).Value
                volume = Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("G" & II).Value
            
            Else ' we are still in the same ticker so just capture the newest close price and add
                 ' the volume
                close_Price = Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("F" & II).Value
                volume = volume + Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("G" & II).Value
            End If
            
        Next II ' loop thru the rows in the worksheet
'       MsgBox ActiveWorkbook.Worksheets(I).Name & " " & NumRows & " " & II & " " & stock_Ticker
'       MsgBox ActiveWorkbook.Worksheets(I).Name
    
    Next I  ' Loop thru Worksheets
End Sub

