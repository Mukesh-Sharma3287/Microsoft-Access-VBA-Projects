VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Tire_Inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdDelete_Click()
     Call ModSave.DeleteTireInventoryData
End Sub

Private Sub cmdReset_Click()
     Call ModSave.ResetTireInventory
End Sub

Private Sub cmdSave_Click()
    Call ModSave.fnTireInventorySave("Save")
End Sub

Private Sub cmdUpdate_Click()
     Call ModSave.fnTireInventorySave("Update")
End Sub

Private Sub txtTireID_AfterUpdate()
    If Form_Tire_Inventory.txtTireID.Value <> "" Then
        Call ModSave.fnSummaryQuery(Form_Tire_Inventory.txtTireID.Value)
    End If
End Sub

