Attribute VB_Name = "Module1"
Option Explicit

Public bWORKING As Boolean
Public bKEEPWORKING As Boolean

Sub deja_vu()
    'never let it run on top of itself
    If bWORKING Then Exit Sub
    bWORKING = True

    'do something here; refresh connections or whatever

    Debug.Print Now 'just to show it did something

    If bKEEPWORKING Then _
        Application.OnTime Now + TimeSerial(0, 0, 5), "deja_vu"
    bWORKING = False
End Sub

