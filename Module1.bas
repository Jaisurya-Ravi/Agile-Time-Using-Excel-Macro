Attribute VB_Name = "Module1"
Public selectedSheetName As String

Sub ShowSheetSelectorForMail()
    ' Show the UserForm to select a sheet
    selectedSheetName = ""
    SheetSelector.Show

    ' Check if a sheet was selected
    If selectedSheetName <> "" Then
        MsgBox "You selected: " & selectedSheetName & "      **** Are you sure want to contine? "
        ' Call SendMailFromWorkBook with the selected sheet name
        ThisWorkbook.SendMailFromWorkBook selectedSheetName
    Else
        MsgBox "No sheet selected.", vbExclamation
    End If
End Sub


Sub ShowSheetSelectorForUploadTimeSheetDetails()
    ' Show the UserForm to select a sheet
    selectedSheetName = ""
    SheetSelector.Show

    ' Check if a sheet was selected
    If selectedSheetName <> "" Then
        MsgBox "You selected: " & selectedSheetName
        ' Call SendMailFromWorkBook with the selected sheet name
        ThisWorkbook.BasicSeleniumVBAIntegration selectedSheetName
    Else
        MsgBox "No sheet selected.", vbExclamation
    End If
End Sub

