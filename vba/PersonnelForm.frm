' VBA UserForm Logic for Personnel Entry ' Fields: Name, Qualification, Dept, HireDate

Private Sub btnSave_Click() Dim wsData As Worksheet: Set wsData = ThisWorkbook.Sheets("Personnel_Data") Dim lastRow As Long: lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row + 1

' Data Validation Logic
If Me.txtName.Value = "" Or Me.txtQual.Value = "" Then
    MsgBox "Please fill all required fields.", vbExclamation
    Exit Sub
End If

' Commit to Database
wsData.Cells(lastRow, 1).Value = Me.txtName.Value
wsData.Cells(lastRow, 2).Value = Me.txtQual.Value
wsData.Cells(lastRow, 3).Value = Me.cmbDept.Value
wsData.Cells(lastRow, 4).Value = Date

MsgBox "Personnel record added successfully.", vbInformation
Unload Me


End Sub