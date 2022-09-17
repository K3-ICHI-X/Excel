VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USDoc 
   Caption         =   "Add Doctor"
   ClientHeight    =   1635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3345
   OleObjectBlob   =   "USDoc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

Sheet6.Activate

emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1

Cells(emptyRow, 1).Value = TextBox2.Value
Cells(emptyRow, 2).Value = TextBox1.Value

TextBox1 = ""
TextBox2 = ""
Call Dgetdata
Sheet1.Activate

Unload Me

End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub



