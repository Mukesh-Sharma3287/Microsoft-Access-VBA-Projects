VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SubPartInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
    
    On Error Resume Next
        Form_Part_Tire_Inventory.lblID.Caption = Form_SubPartInventory.ID.Value
        Form_Part_Tire_Inventory.txtDBRPartNo.Value = Form_SubPartInventory.DBR_Part_Number.Value
        Form_Part_Tire_Inventory.txtMFGPartNo.Value = Form_SubPartInventory.MFG_Part_Number.Value
        Form_Part_Tire_Inventory.txtCarType.Value = Form_SubPartInventory.Car_Type.Value
        Form_Part_Tire_Inventory.txtPartType.Value = Form_SubPartInventory.Part_Type.Value
        Form_Part_Tire_Inventory.txtBrand.Value = Form_SubPartInventory.Brand.Value
        Form_Part_Tire_Inventory.txtModel.Value = Form_SubPartInventory.Model.Value
        Form_Part_Tire_Inventory.txtDescription.Value = Form_SubPartInventory.Description.Value
        Form_Part_Tire_Inventory.txtDataAcquired.Value = Form_SubPartInventory.Date_Acquired.Value
        Form_Part_Tire_Inventory.txtDataServiced.Value = Form_SubPartInventory.Date_Serviced.Value
        Form_Part_Tire_Inventory.txtVendor.Value = Form_SubPartInventory.Vendor.Value
        Form_Part_Tire_Inventory.txtPurchaseInvoice.Value = Form_SubPartInventory.Purchase_Invoice.Value
        Form_Part_Tire_Inventory.txtCost.Value = Form_SubPartInventory.Cost.Value
        Form_Part_Tire_Inventory.txtCondPurchased.Value = Form_SubPartInventory.Condition_Purchased.Value
        Form_Part_Tire_Inventory.txtCurrentCond.Value = Form_SubPartInventory.Current_Condition.Value
        Form_Part_Tire_Inventory.ddnLocation.Value = Form_SubPartInventory.Location.Value
        Form_Part_Tire_Inventory.ddnActiv.Value = Form_SubPartInventory.Active.Value
        Form_Part_Tire_Inventory.txtNote.Value = Form_SubPartInventory.Note.Value
        Form_Part_Tire_Inventory.txtPath.Value = Form_SubPartInventory.ImgPath.Value
        Form_Part_Tire_Inventory.img.Picture = Form_SubPartInventory.ImgPath.Value
    On Error GoTo 0

End Sub
