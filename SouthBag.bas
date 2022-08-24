Attribute VB_Name = "Southbag"
Sub MainSouth()
    Dim ValueSouth As Single
    
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
        ValueSouth = Sheets("Sheet2").Range("C3")

         UpdateProgressBar ValueSouth

End Sub


Sub UpdateProgressBar(ValueSouth As Single)
    With South
Application.WindowState = xlMaximized
' Update the Caption property of the Frame control.
        .LabelSouth.Caption = Format(ValueSouth, "0%")

' Widen the Label control.
        .LabelProgressSouth.Height = 300 - ValueSouth * 300

    End With
 

End Sub

