VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SubEquipmentMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
    
   'On Error Resume Next
   Form_Equipment_Maintenence.lblID.Caption = Form_SubEquipmentMaintenance.ID.Value
   Form_Equipment_Maintenence.txtServiceRecord.Value = Form_SubEquipmentMaintenance.Service_Record.Value
   Form_Equipment_Maintenence.txtDate.Value = Form_SubEquipmentMaintenance.Service_Date.Value
   Form_Equipment_Maintenence.txtEquipment.Value = Form_SubEquipmentMaintenance.Equipment.Value
   Form_Equipment_Maintenence.txtServiceType.Value = Form_SubEquipmentMaintenance.Service_Type.Value
   Form_Equipment_Maintenence.txtServiceBy.Value = Form_SubEquipmentMaintenance.Service_By.Value
   Form_Equipment_Maintenence.txtDescription.Value = Form_SubEquipmentMaintenance.Description.Value
   Form_Equipment_Maintenence.txtNote.Value = Form_SubEquipmentMaintenance.Notes.Value
   'On Error GoTo 0
   
End Sub

    
