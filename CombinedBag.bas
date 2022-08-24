Attribute VB_Name = "CombinedBag"
Sub Main()
    Dim PctDone As Single
    Dim Value1 As Single
    Dim Value2 As Single
    
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
        Value1 = Sheets("Sheet2").Range("C2")
        Value2 = Sheets("Sheet2").Range("C3")
         UpdateProgressBar PctDone

End Sub


Sub UpdateProgressBar(PctDone As Single)
    With Combined
Application.WindowState = xlMaximized
' Update the Caption property of the Frame control.
        .Label2.Caption = Format(PctDone, "0%")
        .Label3.Caption = Format(Sheets("Sheet2").Range("C2"), "0%")
        .Label4.Caption = Format(Sheets("Sheet2").Range("C3"), "0%")
' Widen the Label control.
        .LabelProgress.Height = 300 - PctDone * 300
         .LabelProgress2.Height = 300 - Sheets("Sheet2").Range("C2") * 300
         .LabelProgress3.Height = 300 - Sheets("Sheet2").Range("C3") * 300
    End With
 


 

 
End Sub


