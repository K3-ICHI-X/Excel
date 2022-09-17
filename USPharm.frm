VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USPharm 
   Caption         =   "Add Dispenser"
   ClientHeight    =   1650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3480
   OleObjectBlob   =   "USPharm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USPharm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

Sheet5.Activate

emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1

Cells(emptyRow, 1).Value = TextBox2.Value
Cells(emptyRow, 2).Value = TextBox1.Value

Pgetdata

Sheet1.Activate

Unload Me

End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
