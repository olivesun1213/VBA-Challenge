Attribute VB_Name = "Module1"
Sub Runonallsheet()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call codeforonesheet
    Next
    Application.ScreenUpdating = True
End Sub

Sub codeforonesheet()
On Error Resume Next

'
    'create a Ticker Column
Range("O1") = "Ticker"
Dim A As Object, c As Variant, j As Long, LR1 As Long
Set A = CreateObject("Scripting.Dictionary")
LRA = Cells(Rows.Count, "A").End(xlUp).Row
c = Range("A2:A" & LRA)
For t = 1 To UBound(c, 1)
  A(c(t, 1)) = 1
Next t
Range("O2").Resize(A.Count) = Application.Transpose(A.keys)




'create Unique ID that will be used in VLOOKUP later
Dim ingLastrow As Long
    
    Range("I1") = "Uniqe ID"
    
   Range("I2:I" & LRA).Formula = "=A2 & """" & B2"


'find Max and Min number in date column so we know the open and close date

LRB = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
Cells(2, 24).Value = WorksheetFunction.Max(Range("b2:b" & LRB))
Cells(3, 24).Value = WorksheetFunction.Min(Range("b2:b" & LRB))

'coopy column C-G to create a new table for vlookup
Columns("C:G").Select
    Selection.Copy
    Columns("J:J").Select
    ActiveSheet.Paste
    
'define row and column for outputs

ClosePrice_row = Range("y2").Row
ClosePrice_clm = Range("y2").Column

OpenPrice_row = Range("z2").Row
OpenPrice_clm = Range("z2").Column

totalvolume_row = Range("r2").Row
Totalvolume_clm = Range("r2").Column


PriceChange_row = Range("p2").Row
PriceChange_clm = Range("p2").Column

PcntChange_row = Range("q2").Row
PcntChange_clm = Range("q2").Column

Range("y1") = "OpenPrice"
Range("z1") = "ClosePrice"
Range("p1") = "PriceChange"
Range("q1") = "PercentageChange"
Range("r1") = "TotalVolume"
 
Dim LRO As Long

LRO = ActiveSheet.Range("o" & Rows.Count).End(xlUp).Row

 For i = 2 To LRO
 
    
    
     var1 = Cells(i, 15).Value
     var2 = Cells(2, 24).Value
     var3 = Cells(3, 24).Value
     
     Cells(ClosePrice_row, ClosePrice_clm) = Application.WorksheetFunction.VLookup(var1 & var2, Range("i2:m" & LRB), 5)    'GetClosing Price
     
     Cells(OpenPrice_row, OpenPrice_clm) = Application.WorksheetFunction.VLookup(var1 & var3, Range("i2:m" & LRB), 2)    'GetOpenning price
     
     Cells(PriceChange_row, PriceChange_clm) = Cells(ClosePrice_row, ClosePrice_clm) - Cells(OpenPrice_row, OpenPrice_clm)   'Cal PriceChange
     
     Cells(PcntChange_row, PcntChange_clm) = Cells(PriceChange_row, PriceChange_clm) / Cells(OpenPrice_row, OpenPrice_clm)   'Cal PercentChange
     
     Cells(totalvolume_row, Totalvolume_clm) = Application.WorksheetFunction.SumIf(Range("a2:a" & LRB), var1, Range("n2:n" & LRB))    'Sum Total Volume
     
     
     
     ClosePrice_row = ClosePrice_row + 1
    
     OpenPrice_row = OpenPrice_row + 1
     
     PriceChange_row = PriceChange_row + 1
    
     PcntChange_row = PcntChange_row + 1

     totalvolume_row = totalvolume_row + 1
     
     Next i
     
    
  
  ' change percent result to %
'
    Columns("Q:Q").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
  
'color change based on yearly change value
Dim LRP As Long
LRP = ActiveSheet.Range("P" & Rows.Count).End(xlUp).Row

Dim r1 As Range

   For j = 2 To LRP
      Set r1 = Range("P" & j)
      If r1.Value <= 0 Then r1.Interior.Color = vbRed
      If r1.Value > 0 Then r1.Interior.Color = vbGreen
     
   Next j
 
 
   'find greatest increaset and name & value
   Cells(2, 20) = "Greatest % Increase"
   Cells(2, 21).Value = WorksheetFunction.Max(Range("Q2:Q" & LRP))
   
   'find greatest decrease and display name & value
   Cells(3, 20) = "Greatest % Decrease"
   Cells(3, 21).Value = WorksheetFunction.Min(Range("Q2:Q" & LRP))
   
   'find max total volume and display  name & value
   Cells(4, 20) = "Greatest Total Volume"
   Cells(4, 21).Value = WorksheetFunction.Max(Range("R2:R" & LRP))
   
   'change Greatest % increase & decrease to % format
   Cells(2, 21).Select
   Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
   
   Cells(3, 21).Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"

End Sub



