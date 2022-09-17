VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmLbPrinter 
   Caption         =   "Label Printer"
   ClientHeight    =   5100
   ClientLeft      =   4050
   ClientTop       =   4380
   ClientWidth     =   10320
   OleObjectBlob   =   "FrmLbPrinter.frx":0000
End
Attribute VB_Name = "FrmLbPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCaseNumber_Change()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

If WorksheetFunction.CountIf(Sheet2.Range("A:A"), Me.cbCaseNumber.Value) = 0 Then

Me.cbWard = ""
Me.txtName = ""
Me.cbDoctor = ""
Me.txID = ""

Exit Sub

End If

With Me

On Error Resume Next
Me.cbWard = Application.WorksheetFunction.VLookup(CDbl(Me.cbCaseNumber), Sheet2.Range("A:E"), 4, 0)
Me.txtName = Application.WorksheetFunction.VLookup(CDbl(Me.cbCaseNumber), Sheet2.Range("A:E"), 2, 0)
Me.cbDoctor = Application.WorksheetFunction.VLookup(CDbl(Me.cbCaseNumber), Sheet2.Range("A:E"), 5, 0)
Me.txID = Application.WorksheetFunction.VLookup(CDbl(Me.cbCaseNumber), Sheet2.Range("A:E"), 3, 0)

If Err <> 0 Then

Me.cbWard = Application.WorksheetFunction.VLookup(Me.cbCaseNumber, Sheet2.Range("A:E"), 4, 0)
Me.txtName = Application.WorksheetFunction.VLookup(Me.cbCaseNumber, Sheet2.Range("A:E"), 2, 0)
Me.cbDoctor = Application.WorksheetFunction.VLookup(Me.cbCaseNumber, Sheet2.Range("A:E"), 5, 0)
Me.txID = Application.WorksheetFunction.VLookup(Me.cbCaseNumber, Sheet2.Range("A:E"), 3, 0)

If Err <> 0 Then

Me.cbWard = ""
Me.txtName = ""
Me.cbDoctor = ""
Me.txID = ""

End If
End If
End With

End Sub


Private Sub cbAbout_Click()
MsgBox "Offline Label Form Version 1.7.2 alpha" & vbNewLine & " " & vbNewLine & "" & vbNewLine & "Created By Earl Borcherds"
End Sub

Private Sub cbClose_Click()
Unload Me
End Sub

Private Sub cbMaterial_Change()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

If WorksheetFunction.CountIf(Sheet8.Range("B:B"), Me.cbMaterial.Value) = 0 Then

Me.lbSAP = ""
Me.lbUOM = ""

End If

If Me.cbMaterial <> "" Then
 On Error Resume Next
Me.lbSAP = Application.WorksheetFunction.VLookup(cbMaterial, Sheet8.Range("b:e"), 3, False)
Me.lbUOM = Application.WorksheetFunction.VLookup(cbMaterial, Sheet8.Range("b:e"), 2, 0)
 If Err <> 0 Then
    Me.lbSAP = ""
    Me.lbUOM = ""
    End If
End If


End Sub


Private Sub cmdbAddWard_Click()
USWard.Show
End Sub

Private Sub cmdbClose_Click()

    Unload Me
    
End Sub


Private Sub cmdbPrint_Click()

    Dim intCopies, intRegRow As Integer
    Dim dtAdmDate As Date
    Dim dtAdmTime As Date
    Dim strDate, strTime, strWard, strQty, strMaterial, strInstructions, strPharmacist, strCaseNumber, strtxtName, strDoctor, strSAP, strUOM As String

    strMaterial = Me.cbMaterial.Text
    strQty = Me.txtQty.Text
    strPharmacist = Me.cbDispenser.Text
    strDate = Format(Me.lbDate, "YYYY/MM/dd")
    strTime = Format(Me.lbTime, "HH:mm")
    strWard = StrConv(Me.cbWard, vbProperCase)
    strCaseNumber = Me.cbCaseNumber.Text
    strtxtName = Me.txtName.Text
    intCopies = Me.txtCopies
    strInstructions = Me.txtInstructions.Text
    strDoctor = Me.cbDoctor.Text
    strSAP = Me.lbSAP
    strUOM = Me.lbUOM

    ThisWorkbook.Names("MedName").RefersToRange = strMaterial
    ThisWorkbook.Names("MedQty").RefersToRange = strQty
    ThisWorkbook.Names("Directions").RefersToRange = strInstructions
    ThisWorkbook.Names("Ward").RefersToRange = strWard
    ThisWorkbook.Names("Dispenser").RefersToRange = strPharmacist
    ThisWorkbook.Names("ScriptDate").RefersToRange = strDate
    ThisWorkbook.Names("ScriptTime").RefersToRange = strTime
    ThisWorkbook.Names("CaseNumber").RefersToRange = strCaseNumber
    ThisWorkbook.Names("Name").RefersToRange = strtxtName
    ThisWorkbook.Names("Doctor").RefersToRange = strDoctor
    ThisWorkbook.Names("UOM").RefersToRange = strUOM

    Sheet1.Activate

    lbDate = Format(Now, "DD/MM/YYYY")
    lbTime = Format(Now, "HH:mm")

'Error Messages
      If cbCaseNumber.Text = "" Then
        cbCaseNumber.SetFocus
        Beep
        MsgBox "Please Select or Input Case Number."
            Exit Sub
    End If
   
    If txtName.Text = "" Then
        txtName.SetFocus
        Beep
        MsgBox "Please input patient name."
            Exit Sub
    End If
    
     If cbWard.Text = "" Then
        FrmLbPrinter.cbWard.SetFocus
        Beep
        MsgBox "Please Select Ward."
            Exit Sub
    End If
    
        If cbDoctor.Text = "" Then
        FrmLbPrinter.cbDoctor.SetFocus
        Beep
        MsgBox "Please Select Doctor."
            Exit Sub
    End If

    If cbMaterial.Text = "" Then
        FrmLbPrinter.cbMaterial.SetFocus
        Beep
        MsgBox "Please Select Material."
            Exit Sub
    End If

    If txtQty.Text = "" Then
        txtQty.SetFocus
        Beep
        MsgBox "Please Select Quantity."
            Exit Sub
    End If

    If txtInstructions.Text = "" Then
        txtInstructions.SetFocus
        Beep
        MsgBox "Please input instructions."
            Exit Sub
    End If

    If cbDispenser.Text = "" Then
        cbDispenser.SetFocus
        Beep
        MsgBox "Please select dispenser."
            Exit Sub
    End If

    If txtCopies.Text = "" Then
        txtCopies.SetFocus
        Beep
        MsgBox "Please Select Number Label Copies."
            Exit Sub
    End If

    If txtCopies = "0" Then
        txtCopies.SetFocus
        Beep
        MsgBox "Please Select Quantity Greater than Zero."
            Exit Sub
    End If

    If Not (CurrentPrinterLabel Like "*ZDesigner*") Then
        MsgBox "Please Select a 'ZDesigner' Label Printer."
            Beep
        Application.Dialogs(xlDialogPrinterSetup).Show
        CurrentPrinterLabel.Caption = Application.ActivePrinter

        Exit Sub
    End If
    
   
'Print
'Sheet11.Activate

If EnAbbr.Value = True Then Call ac

'Sheet7.Activate
    Sheet7.Visible = True
    Sheet7.PrintOut Copies:=intCopies
    Sheet7.Visible = False
    
If ADLAB.Value = True Then
    Sheet10.Visible = True
    Sheet10.PrintOut
    Sheet10.Visible = False
    
End If

'Save data to Sheets
    Sheet3.Activate
    emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1

    Cells(emptyRow, 1).Value = cbCaseNumber.Value
    Cells(emptyRow, 2).Value = txtName.Value
    Cells(emptyRow, 3).Value = cbWard.Value
    Cells(emptyRow, 4).Value = lbSAP
    Cells(emptyRow, 5).Value = cbMaterial.Value
    Cells(emptyRow, 6).Value = txtQty.Value
    Cells(emptyRow, 7).Value = txtInstructions.Value
    Cells(emptyRow, 8).Value = cbDoctor.Value
    Cells(emptyRow, 9).Value = cbDispenser.Value
    Cells(emptyRow, 10).Value = lbDate
    Cells(emptyRow, 11).Value = lbTime

    cbMaterial = ""
    txtInstructions = ""
    ADLAB.Value = False

    Sheet2.Activate

    If WorksheetFunction.CountIf(Sheet2.Range("A:A"), Me.cbCaseNumber.Value) = 0 Then

        emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1

        Sheet2.Cells(emptyRow, 1).Value = cbCaseNumber.Value
        Sheet2.Cells(emptyRow, 2).Value = txtName.Value
        Sheet2.Cells(emptyRow, 4).Value = cbWard.Value
        Sheet2.Cells(emptyRow, 5).Value = cbDoctor.Value

        Sheet1.Activate

        Call getdata

    End If
Sheet3.Activate

End Sub


Private Sub cmdbCPrint_Click()

    Application.Dialogs(xlDialogPrinterSetup).Show
    CurrentPrinterLabel.Caption = Application.ActivePrinter

End Sub

Private Sub cmdbAddDoc_Click()
    USDoc.Show
End Sub

Private Sub cmdbAddDis_Click()
    USPharm.Show
End Sub


Private Sub txtCopies_Change()

    With Me.txtCopies
        If .Text Like "[!0-9]" Or Val(.Text) < -1 Or .Text Like "?*[!0-9]*" Then
            Beep
            MsgBox "Input a valid Number."
            .Text = Left(.Text, Len(.Text) - 1)
        End If
    End With

End Sub


Private Sub txtQty_Change()

    With Me.txtQty
        If .Text Like "[!0-9]" Or Val(.Text) < -1 Or .Text Like "?*[!0-9]*" Then
            Beep
            MsgBox "Input a valid Number."
            .Text = Left(.Text, Len(.Text) - 1)
        End If
    End With

End Sub

Private Sub txtSAP_AfterUpdate()

    If WorksheetFunction.CountIf(Sheet5.Range("A:A"), Me.txtSAP.Value) = 0 Then

        Me.cmbbMaterial = ""
        Me.UOM = ""

    End If

    If Me.txtSAP <> "" Then
        On Error Resume Next
        Me.cmbbMaterial = Application.WorksheetFunction.VLookup(Me.txtSAP, Sheet5.Range("A:e"), 2, False)
        Me.UOM = Application.WorksheetFunction.VLookup(Me.txtSAP, Sheet5.Range("a:e"), 3, 0)
        If Err <> 0 Then
            Me.cmbbMaterial = ""
            Me.UOM = ""
        End If
    End If

End Sub

Private Sub UserForm_Activate()
Call getdata
    CurrentPrinterLabel.Caption = Application.ActivePrinter


    'With Range("Reg_CNum[Case Number]")
    '.NumberFormat = "General"
  '.Value = .Value
 'End With
    
End Sub
