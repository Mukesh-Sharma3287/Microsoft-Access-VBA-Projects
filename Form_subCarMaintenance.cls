VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subCarMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
    On Error Resume Next
    Call SetNewId
    Form_Car_Maintenence.lblID.Caption = Form_subCarMaintenance.Service_Record.Value
    Form_Car_Maintenence.txtDate.Value = Form_subCarMaintenance.Start_Date.Value
    Form_Car_Maintenence.txtCar.Value = Form_subCarMaintenance.Car.Value
    Form_Car_Maintenence.txtServiceType.Value = Form_subCarMaintenance.Service_Type.Value
    Form_Car_Maintenence.txtServiceBy.Value = Form_subCarMaintenance.Service_By.Value
    On Error GoTo 0
End Sub

