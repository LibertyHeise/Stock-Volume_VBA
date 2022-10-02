VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stocks()
' Dimension variables
Dim total As Double
'Create title row

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'initalize variables
total_volume = 0
c_pointer = 2
i_pointer = 2
'Determine number of rows
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'Start loop
For i = 2 To RowCount
    If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then

'Define what will be called in the looping process

        total_volume = total_volume + Cells(i, "G").Value
        startingPrice = Cells(i, "C").Value
        closingPrice = Cells(i, "F").Value
        YearlyChange = closingPrice - startingPrice
        PercentChange = YearlyChange / startingPrice * 100
        
'Using cell pointers determined by their initial value and place, determine what happens next
        
        Cells(i_pointer, "I").Value = Cells(i, "A").Value
        Cells(i_pointer, "J").Value = YearlyChange
        Cells(i_pointer, "K").Value = "%" & PercentChange
        Cells(i_pointer, "L").Value = total_volume
        
'Using the change, negative or positive, make the box turn red or green
'Red,  Green, no color
        If YearlyChange < 0 Then
            Cells(i_pointer, "J").Interior.ColorIndex = 3
        ElseIf YearlyChange > 0 Then
            Cells(i_pointer, "J").Interior.ColorIndex = 4
        Else
            Cells(i_pointer, "J").Interior.ColorIndex = 2
        End If

'Reset and redirect for the next iteration

        c_pointer = i + 1
        i_pointer = i_pointer + 1
        total_volume = 0
        


    Else
        total_volume = total_volume + Cells(i, "G").Value
    End If
  
Next i
End Sub



