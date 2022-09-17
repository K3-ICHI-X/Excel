Attribute VB_Name = "abbreviationList"
Public Sub ac()

Dim i As Integer

For i = 2 To 49

Sheet11.Activate

Range("D2").Select

Selection.Replace What:=Cells(i, 1).Value, Replacement:=Cells(i, 2).Value, lookat:=xlPart, searchorder:=xlByRows, MatchCase:=False

Next

Sheets("Subs").Cells(1, 1).Select

End Sub
