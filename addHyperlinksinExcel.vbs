Sub AddHyperlinkToCell()
Dim foundRng As Range
endCount = Worksheets("Sheet1").UsedRange.Rows.Count
For i = 1 To endCount
  Rem find the word in Shee1 - col A in Sheet 2 range.
  Set Cell = Worksheets("Sheet2").Range("A1:A101").Find(Range("A" & i).Value)

  addr = Trim(Worksheets("Sheet2").Range("B" & Cell.Row).Value)
  Debug.Print addr
  ActiveSheet.Hyperlinks.Add Range("A" & i), Address:=addr
  Next i
End Sub


