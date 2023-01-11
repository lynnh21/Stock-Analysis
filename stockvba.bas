VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_analysis()

'----------------------FIRST PART--------------------

    [i1] = "Ticker"
    [j1] = "Yearly Change"
    [k1] = "Percent Change"
    [l1] = "Total Stock Volume"
   
    tIndex = 2
    lowT = ""
    lowP = 0
    lowV = 0
    openValue = 0
    total_volume = 0
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
   
    For i = 2 To lastRow
   
        total_volume = total_volume + Cells(i, "G").Value
   
        If openValue = 0 Then
            openValue = Cells(i, "C").Value
        End If
       
        If Cells(i, 1) <> Cells(i + 1, 1) Then
            Cells(tIndex, 9).Value = Cells(i, 1).Value

            yearly_change = Cells(i, "F").Value - openValue
            
            
            
            Cells(tIndex, "J").Value = yearly_change
           
            If yearly_change < 0 Then
                Cells(tIndex, "J").Interior.ColorIndex = 3
            Else
                Cells(tIndex, "J").Interior.ColorIndex = 4
            End If
           
            Cells(tIndex, "L").Value = total_volume
           percentage_change = FormatPercent(yearly_change / openValue)
            Cells(tIndex, "K").Value = percentage_change
           
           If yearly_change < lowV Then
                lowT = Cells(i, "A")
                lowV = yearly_change
                lowP = percentage_change
            End If
            
            total_volume = 0
            openValue = 0
            tIndex = tIndex + 1
        End If
       
   
    Next i
Range("P3") = lowT
Range("Q3") = lowP

 '---------------------BONUS PART-------------------------

Range("O2") = "Greatest % increase"
Range("O3") = "Lowest % decrease"
Range("O4") = "Greatest total Volume"
Range("P1") = "Ticker"
Range("Q1") = "Value"

'-------------------------------------------------------
lastrow2 = Cells(Rows.Count, "I").End(xlUp).Row
greatest_percent = 0
greatest_percent_index = 2

For j = 2 To lastrow2
    If Cells(j, "K") > greatest_percent Then
        greatest_percent = Cells(j, "K").Value
        greatest_percent_index = j
    Else: greatest_percent = greatest_percent
        greatest_percent_index = greatest_percent_index
    End If
Next j

Cells(2, 16).Value = Cells(greatest_percent_index, 9).Value
Cells(2, 17).Value = FormatPercent(greatest_percent)

'-------------------------------------------------------

highest_volume = 0
highest_volume_index = 2
For k = 2 To lastrow2
    If Cells(k, 12).Value > highest_volume Then
    highest_volume = Cells(k, 12).Value
        highest_volume_index = k
    Else: highest_volume = highest_volume
        highest_volume_index = highest_volume_index
    End If
Next k

Cells(4, 16).Value = Cells(highest_volume_index, 9).Value
Cells(4, 17).Value = highest_volume
    
    
    
End Sub



