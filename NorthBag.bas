Attribute VB_Name = "NorthBag"
Sub MainNorth()
    Dim PctDone As Single
    
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
        PctDone = Sheets("Sheet2").Range("C1")

         UpdateProgressBar PctDone

End Sub


Sub UpdateProgressBar(PctDone As Single)
    With North
Application.WindowState = xlMaximized
' Update the Caption property of the Frame control.
        .LabelNorth.Caption = Format(PctDone, "0%")

' Widen the Label control.
        .LabelProgressNorth.Height = 300 - PctDone * 300

    End With
 


 

 
End Sub
