VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Equipment_Maintenence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdDelete_Click()
     Call ModSave.DeleteEquipmentMaintenance
End Sub

Private Sub cmdReset_Click()
     Call ModSave.ResetEquipmentMaintenance
End Sub

Private Sub cmdSave_Click()
    Call ModSave.fnEquipmentMaintenanceSave("Save")
End Sub

Private Sub cmdUpdate_Click()
     Call ModSave.fnEquipmentMaintenanceSave("Update")
End Sub
