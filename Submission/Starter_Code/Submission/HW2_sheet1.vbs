VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_analysis():
    ' Set dimensions
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowcount As Long
    Dim perchange As Double
    Dim days As Integer
    Dim daychange As Double
    Dim avechange As Double
    ' Title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    ' Initial values
    j = 0
    total = 0
    change = 0
    start = 2
    ' Row number of last row
    rowcount = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To rowcount
        ' If ticker changes, print results
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ' Stores results in variables
            total = total + Cells(i, 7).Value
            ' Zero fix
            If total = 0 Then
                ' Print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0
            Else
                ' First non-zero starting value
                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                     Next find_value
                End If
                ' Calculate Change
                change = (Cells(i, 6) - Cells(start, 3))
                perchange = change / Cells(start, 3)
                ' Start of the next stock ticker
                start = i + 1
                ' Print results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = change
                Range("J" & 2 + j).NumberFormat = "0.00"
                Range("K" & 2 + j).Value = perchange
                Range("K" & 2 + j).NumberFormat = "0.00%"
                Range("L" & 2 + j).Value = total
                ' Colors positives green and negatives red
                Select Case change
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select
            End If
            ' Reset for new ticker
            total = 0
            change = 0
            j = j + 1
            days = 0
        ' If same, add results
        Else
            total = total + Cells(i, 7).Value
        End If
    Next i
    ' Max. and min.
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowcount)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowcount)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowcount))
    ' Disregards header row; one less
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowcount)), Range("K2:K" & rowcount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowcount)), Range("K2:K" & rowcount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowcount)), Range("L2:L" & rowcount), 0)
    ' Final ticker for total, greatest percent increase and decrease, average
    Range("P2") = Cells(increase_number + 1, 9)
    Range("P3") = Cells(decrease_number + 1, 9)
    Range("P4") = Cells(volume_number + 1, 9)
End Sub











