Attribute VB_Name = "Data"
Option Explicit

Dim iCaseNumber, intWard, iDoctor, iMaterial, iPharmacist As Integer
Dim strCaseNumber, strWard, strDoctor, strMaterial, strPharmacist As String
Dim dtDate, dtTime As Date
Public Sub getdata()

    dtDate = Format(Now, "DD/MM/YYYY")
    FrmLbPrinter.lbDate = dtDate

    dtTime = Format(Now, "HH:mm")
    FrmLbPrinter.lbTime = Format(Now, "HH:mm")

iCaseNumber = 3
intWard = 1
iDoctor = 1
iMaterial = 2
iPharmacist = 1

FrmLbPrinter.cbCaseNumber.Clear
FrmLbPrinter.cbDoctor.Clear
FrmLbPrinter.cbWard.Clear
FrmLbPrinter.cbMaterial.Clear
FrmLbPrinter.cbDispenser.Clear

    With FrmLbPrinter.cbCaseNumber
        Do Until IsEmpty(Sheet2.Cells(iCaseNumber, 1))
            strCaseNumber = Sheet2.Cells(iCaseNumber, 1)
            .AddItem strCaseNumber
            iCaseNumber = iCaseNumber + 1
        Loop
    End With

    With FrmLbPrinter.cbWard
        Do Until IsEmpty(Sheet4.Cells(intWard, 1))
            strWard = Sheet4.Cells(intWard, 1)
            .AddItem strWard
            intWard = intWard + 1
        Loop
    End With

    With FrmLbPrinter.cbDoctor
       Do Until IsEmpty(Sheet6.Cells(iDoctor, 1))
          strDoctor = Sheet6.Cells(iDoctor, 1) & ", " & Sheet6.Cells(iDoctor, 2)
        .AddItem strDoctor
       iDoctor = iDoctor + 1
    Loop
    End With
     
    With FrmLbPrinter.cbMaterial
        Do Until IsEmpty(Sheet8.Cells(iMaterial, 2))
            strMaterial = Sheet8.Cells(iMaterial, 2)
            .AddItem strMaterial
            iMaterial = iMaterial + 1
        Loop
    End With

    With FrmLbPrinter.cbDispenser
        Do Until IsEmpty(Sheet5.Cells(iPharmacist, 1))
            strPharmacist = Sheet5.Cells(iPharmacist, 1) & ", " & Sheet5.Cells(iPharmacist, 2)
            .AddItem strPharmacist
            iPharmacist = iPharmacist + 1
        Loop
    End With
    
   
             
End Sub

Public Sub Dgetdata()

iDoctor = 1

FrmLbPrinter.cbDoctor.Clear

    
    With FrmLbPrinter.cbDoctor
       Do Until IsEmpty(Sheet6.Cells(iDoctor, 1))
          strDoctor = Sheet6.Cells(iDoctor, 1) & ", " & Sheet6.Cells(iDoctor, 2)
        .AddItem strDoctor
       iDoctor = iDoctor + 1
    Loop
    End With
          
End Sub

Public Sub Pgetdata()

iPharmacist = 1

FrmLbPrinter.cbDispenser.Clear

    
    With FrmLbPrinter.cbDispenser
        Do Until IsEmpty(Sheet5.Cells(iPharmacist, 1))
            strPharmacist = Sheet5.Cells(iPharmacist, 1) & ", " & Sheet5.Cells(iPharmacist, 2)
            .AddItem strPharmacist
            iPharmacist = iPharmacist + 1
        Loop
    End With
              
 End Sub
 Public Sub Wgetdata()
 intWard = 1
 FrmLbPrinter.cbWard.Clear
    With FrmLbPrinter.cbWard
        Do Until IsEmpty(Sheet4.Cells(intWard, 1))
            strWard = Sheet4.Cells(intWard, 1)
            .AddItem strWard
            intWard = intWard + 1
        Loop
    End With
    
 End Sub

Public Sub RegData()

intWard = 1
iDoctor = 1

    With RegForm.cbDoctor
        Do Until IsEmpty(Sheets("Doctor").Cells(iDoctor, 1))
            strDoctor = Sheets("Doctor").Cells(iDoctor, 1) & ", " & Sheets("Doctor").Cells(iDoctor, 2)
            .AddItem strDoctor
            iDoctor = iDoctor + 1
        Loop
    End With
    
       
    With RegForm.cmWard
        Do Until IsEmpty(Sheets("Wards").Cells(intWard, 1))
            strWard = Sheets("Wards").Cells(intWard, 1)
            .AddItem strWard
            intWard = intWard + 1
        Loop
    End With
    
    End Sub
