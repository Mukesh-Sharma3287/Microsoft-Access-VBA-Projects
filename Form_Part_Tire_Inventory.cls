VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Part_Tire_Inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'''***********************************************************************************************************
''' Procedure Name       : cmdBrowse_Click
''' Description          : To Browse the database path
''' Arguments            : NA
''' Returns              : NA
'''***********************************************************************************************************
'
Private Sub cmdBrowseImg_Click()

    'Declare the Variable
    Dim objFileDialog As FileDialog
    Dim vSelected As Variant

    'Set the Object
    Set objFileDialog = Application.FileDialog(msoFileDialogFilePicker)

    'Filter the files
    With objFileDialog
        .AllowMultiSelect = False
        .ButtonName = "Select File"
        .Filters.Clear
        .Filters.Add "Image File", "*.jpg;*.png", 1
        .Title = "Select Image Path"
        .Show

        For Each vSelected In objFileDialog.SelectedItems
            Form_Part_Tire_Inventory.txtPath.Value = vSelected
        Next

    End With
    
    If InStr(1, Form_Part_Tire_Inventory.txtPath.Value, ".jpg", vbTextCompare) > 0 Then
         Form_Part_Tire_Inventory.img.Visible = True
         Form_Part_Tire_Inventory.img.Picture = Form_Part_Tire_Inventory.txtPath.Value
    End If
    
    'Clear the memory
    Set objFileDialog = Nothing

End Sub


Private Sub cmdDelete_Click()
    Call ModSave.DeletePartInventoryData
End Sub

Private Sub cmdReset_Click()
    Call ModSave.ResetPartInventory
End Sub

Private Sub cmdSave_Click()
     Call ModSave.fnPartTireInventorySave("Save")
End Sub

Private Sub cmdUpdate_Click()
    Call ModSave.fnPartTireInventorySave("Update")
End Sub

Private Sub txtPath_AfterUpdate()

   
End Sub

Private Sub txtPath_Click()

End Sub
