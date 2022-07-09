Sub Proper_Case()
   ' Loop to cycle through each cell in the specified range.
   For Each x In Range("C1:C5")
      ' There is not a Proper function in Visual Basic for Applications.
      ' So, you must use the worksheet function in the following form:
      x.Value = Application.Proper(x.Value)
   Next
End Sub
