Attribute VB_Name = "Module1"
Sub Stock_Data_Homework_Macro()
    ' applying to all worksheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    
        ' Adding text in cells
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change($)"
        ws.Cells(1, 11).Value = "Percent Change(%)"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        ' Define variables
        Dim LastRow As Long
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim Tick As String
        Dim TotalLocation As Long
        TotalLocation = 2
        Dim TTotal As LongLong
        TTotal = 0
        Dim Difference As Double
        Difference = 0
        Dim YearChange As Double
        YearChange = ws.Cells(2, 3).Value
        Dim YearChangeClose As Double
        YearChangeClose = 0
        Dim maxValue As LongLong
        maxValue = -99999

        
        ' Filling column I and L, I as ticker type and L as Total stock value
        ' and filling yearly change + % change columns
        Dim i As Long
        For i = 2 To LastRow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Checking ticker type
                Tick = ws.Cells(i, 1).Value
                ' Adding ticker volume and yearly change
                TTotal = TTotal + ws.Cells(i, 7).Value
                YearChangeClose = ws.Cells(i, 6).Value
                
                
                ' Location of where to print values
                ws.Range("I" & TotalLocation).Value = Tick
                ws.Range("L" & TotalLocation).Value = TTotal
                ws.Range("J" & TotalLocation).Value = YearChangeClose - YearChange
                ws.Range("K" & TotalLocation).Value = (YearChangeClose - YearChange) / YearChange * 100
                
                ws.Cells(4, 16).Value = maxValue
                
                
                TotalLocation = TotalLocation + 1
                ' Reset totals for next ticker type
                YearChange = ws.Cells(i + 1, 3).Value
                TTotal = 0
                YearChangeClose = 0
                
                'changing color of the cells in K if they are +/-
                If Not IsEmpty(ws.Cells(i, 10).Value) Then
                    If ws.Cells(i, 10).Value > 0 Then
                        ws.Cells(i, 10).Interior.ColorIndex = 4
                    ElseIf ws.Cells(i, 10).Value < 0 Then
                        ws.Cells(i, 10).Interior.ColorIndex = 3
                    End If
                End If
            
           
            Else
                ' adding total
                TTotal = TTotal + ws.Cells(i, 7).Value
                                
                ' greatest total vol
                currentValue = ws.Cells(i, 7).Value
                    If currentValue > maxValue Then
                    maxValue = currentValue
                    ws.Cells(4, 15).Value = Tick
                    End If

            End If
            
        Next i
        'finding greatest % increase and decrease
        Dim highVal As Double
        highVal = -99999
        Dim maxValTick As String
        maxValTick = 0
        Dim minVal As Double
        minVal = 999
        Dim minValTick As String
        
        For i = 2 To LastRow
            If ws.Cells(i, 11).Value > highVal Then
                highVal = ws.Cells(i, 11).Value
                maxValTick = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 11).Value < minVal Then
                minVal = ws.Cells(i, 11).Value
                minValTick = ws.Cells(i, 9).Value
            End If
        Next i
        ' prints em
        ws.Cells(2, 16).Value = highVal
        ws.Cells(2, 15).Value = maxValTick
        ws.Cells(3, 16).Value = minVal
        ws.Cells(3, 15).Value = minValTick
        
    'setting the color for cells
    Set Rng = ws.Range("K2:K" & LastRow)
    Rng.Interior.Color = xlNone
        For Each Cell In Rng
            If IsNumeric(Cell.Value) Then
                If Cell.Value > 0 Then
                    Cell.Interior.Color = RGB(0, 255, 0) ' Green
                ElseIf Cell.Value < 0 Then
                    Cell.Interior.Color = RGB(255, 0, 0) ' Red
                End If
            End If
        Next Cell
    ' resetting variables. had to do this, for some reason it broke w/o it
    TTotal = 0
    Difference = 0
    TotalLocation = 2
    YearChange = 0
    YearChangeClose = 0
    maxVal = -99999

    Next ws
    
End Sub
