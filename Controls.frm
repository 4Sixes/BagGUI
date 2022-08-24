VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Controls 
   Caption         =   "Baghouse Monitoring System"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5355
   OleObjectBlob   =   "Controls.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Controls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Combined.show

End Sub

Private Sub CommandButton2_Click()
Combined.Hide
Calculate
Combined.show
End Sub

Private Sub CommandButton5_Click()
West.show
End Sub

Private Sub CommandButton6_Click()
South.show
End Sub

Private Sub CommandButtonExit_Click()


Combined.Hide
North.Hide
South.Hide
West.Hide
Call ShowUserForm



End Sub


Private Sub SystemControls_Click()
North.show
End Sub
