VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SubTireInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
    On Error Resume Next
        Form_Tire_Inventory.lblID.Caption = Form_SubTireInventory.ID.Value
        Form_Tire_Inventory.txtTireID.Value = Form_SubTireInventory.Tire_Id.Value
        Form_Tire_Inventory.txtBarCode.Value = Form_SubTireInventory.Bar_Code.Value
        Form_Tire_Inventory.txtSet.Value = Form_SubTireInventory.Set.Value
        Form_Tire_Inventory.txtManufacturers.Value = Form_SubTireInventory.Manufacturer.Value
        Form_Tire_Inventory.txtType.Value = Form_SubTireInventory.Type.Value
        Form_Tire_Inventory.txtDesignatedCar.Value = Form_SubTireInventory.Designated_Car.Value
        Form_Tire_Inventory.txtModelNumber.Value = Form_SubTireInventory.Model_Number.Value
        Form_Tire_Inventory.txtDescription.Value = Form_SubTireInventory.Description.Value
        Form_Tire_Inventory.txtDataAcquired.Value = Form_SubTireInventory.Date_Acquired.Value
        Form_Tire_Inventory.txtDateMounted.Value = Form_SubTireInventory.Date_Mounted.Value
        Form_Tire_Inventory.txtVendor.Value = Form_SubTireInventory.Vendor.Value
        Form_Tire_Inventory.txtPurchaseInvoice.Value = Form_SubTireInventory.Purchase_Invoice.Value
        Form_Tire_Inventory.txtCost.Value = Form_SubTireInventory.Cost.Value
        Form_Tire_Inventory.txtCondPurchased.Value = Form_SubTireInventory.Condition_Purchased.Value
        Form_Tire_Inventory.txtCurrentCond.Value = Form_SubTireInventory.Current_Condition.Value
        Form_Tire_Inventory.txtMountedRim.Value = Form_SubTireInventory.Mounted_Rim.Value
        Form_Tire_Inventory.ddnLocation.Value = Form_SubTireInventory.Location.Value
        Form_Tire_Inventory.ddnActiv.Value = Form_SubTireInventory.Active.Value
        Form_Tire_Inventory.txtTotalHeatCycle.Value = Form_SubTireInventory.Total_Heat_Cycle.Value
        Form_Tire_Inventory.txtTotalTimeOnTire.Value = Form_SubTireInventory.Total_Time_On_Tire.Value
        Form_Tire_Inventory.txtNote.Value = Form_SubTireInventory.Note.Value
    On Error GoTo 0
End Sub
