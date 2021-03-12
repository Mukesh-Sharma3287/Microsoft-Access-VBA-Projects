VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Reports_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdCarMaintenance_Click()
     DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "tblCarMaintenance", CurrentProject.Path & "\CarMaintenance"
    DoCmd.OutputTo acOutputForm, "Car_Maintenence", acFormatPDF, CurrentProject.Path & "\CarMaintenance.pdf"
End Sub

Private Sub cmdCarSetup_Click()
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "tblCarSetup", CurrentProject.Path & "\CarSetup"
    DoCmd.OutputTo acOutputForm, "Setup", acFormatPDF, CurrentProject.Path & "\CarSetup.pdf"
End Sub

Private Sub cmdEquipMaintenance_Click()
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "tblEquipmentMaintenance", CurrentProject.Path & "\EquipmentMaintenance"
    DoCmd.OutputTo acOutputForm, "Equipment_Maintenence", acFormatPDF, CurrentProject.Path & "\EquipmentMaintenance.pdf"
End Sub

Private Sub cmdPartsInventory_Click()
     DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "tblPartInventory", CurrentProject.Path & "\PartInventory"
     DoCmd.OutputTo acOutputForm, "Part_Tire_Inventory", acFormatPDF, CurrentProject.Path & "\PartInventory.pdf"
End Sub

Private Sub cmdRaceSession_Click()
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "tblRaceSession", CurrentProject.Path & "\RaceData"
    DoCmd.OutputTo acOutputForm, "Race_Session", acFormatPDF, CurrentProject.Path & "\RaceData.pdf"
End Sub

Private Sub cmdTireInventory_Click()
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "qryTireInventory", CurrentProject.Path & "\TireInventory"
    DoCmd.OutputTo acOutputForm, "Tire_Inventory", acFormatPDF, CurrentProject.Path & "\TireInventory.pdf"
End Sub
