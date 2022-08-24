Attribute VB_Name = "WestBag"
Sub MainWest()
    Dim ValueWest As Single
    
Application.ScreenUpdating = True
    ' Initialize variables.
   ' Counter = 1
  '  RowMax = 1000
    'ColMax = 25

' Loop through cells.
   ' For r = 1 To RowMax
      '  For c = 1 To ColMax
      '  '    'Put a random number in a cell
         '   Cells(r, c) = Int(Rnd * 1000)
      '      Counter = Counter + 1
       ' Next c

' Update the percentage completed.
        ValueWest = Sheets("Sheet2").Range("C2")

         UpdateProgressBar ValueWest

End Sub


Sub UpdateProgressBar(ValueWest As Single)
    With West
Application.WindowState = xlMaximized
' Update the Caption property of the Frame control.
        .LabelWest.Caption = Format(ValueWeste, "0%")

' Widen the Label control.
        .LabelProgressWest.Height = 300 - PctDone * 300

    End With
 
End Sub
