VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SheetSelector 
   Caption         =   "Select to continue"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   OleObjectBlob   =   "SheetSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SheetSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub UserForm_Initialize()
    ' Initialize the ComboBox with sheet names
    RefreshComboBox
End Sub

Private Sub RefreshComboBox()
    Dim ws As Worksheet
    
    ' Clear existing items in the ComboBox
    Me.cmbSheets.Clear
    
    ' Populate the ComboBox with current sheet names
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> "Sheet1" And ws.name <> "Sheet2" And ws.name <> "User Details" And ws.name <> "Master Details" Then
            Me.cmbSheets.AddItem ws.name
        End If
    Next ws
End Sub

Private Sub UserForm_Click()
    ' Check if a sheet is selected
    If Me.cmbSheets.ListIndex = -1 Then
        MsgBox "Please select a sheet.", vbExclamation
    Else
        ' Set the selected sheet name to a public variable
        selectedSheetName = Me.cmbSheets.Value
        Me.Hide
    End If
End Sub

Private Sub UserForm_Activate()
    ' Refresh ComboBox when UserForm is activated (shown)
    RefreshComboBox
End Sub



Private Sub btnOK_Click()
    ' Check if a sheet is selected
    If Me.cmbSheets.ListIndex = -1 Then
        MsgBox "Please select a sheet.", vbExclamation
    Else
        ' Set the selected sheet name to a public variable
        selectedSheetName = Me.cmbSheets.Value
        Me.Hide
    End If
End Sub

