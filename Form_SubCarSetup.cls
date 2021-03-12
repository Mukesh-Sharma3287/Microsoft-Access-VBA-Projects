VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SubCarSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
    
    On Error Resume Next
    Form_Setup.lblID.Caption = Form_SubCarSetup.ID.Value
    Form_Setup.txtSetupId.Value = Form_SubCarSetup.Setup_Id.Value
    Form_Setup.txtCarType.Value = Form_SubCarSetup.Car_Type.Value
    Form_Setup.txtTrackSpecific.Value = Form_SubCarSetup.Track_Specific.Value
    Form_Setup.txtTrackCondition.Value = Form_SubCarSetup.Track_Condition.Value
    Form_Setup.txtTempratureRange.Value = DLookup("[Temprature Range]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtType.Value = DLookup("Type", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    'LF
    Form_Setup.txtLFWeight.Value = DLookup("[LF Weight]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtLFRideHeight.Value = DLookup("[LF Ride Height]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtLFCamber.Value = DLookup("[LF Camber]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtLFToe.Value = DLookup("[LF Toe]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtLFShock.Value = DLookup("[LF Shock]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtLFPressureCold.Value = DLookup("[LF Pressure Cold]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtLFPressureHot.Value = DLookup("[LF Pressure Hot]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    'LR
    Form_Setup.txtLRWeight.Value = DLookup("[LR Weight]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtLRRideHeight.Value = DLookup("[LR Ride Height]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtLRCamber.Value = DLookup("[LR Camber]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtLRToe.Value = DLookup("[LR Toe]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtLRShock.Value = DLookup("[LR Shock]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtLRPressureCold.Value = DLookup("[LR Pressure Cold]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtLRPressureHot.Value = DLookup("[LR Pressure Hot]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    'RF
    Form_Setup.txtRFWeight.Value = DLookup("[RF Weight]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtRFRideHeight.Value = DLookup("[RF Ride Height]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtRFCamber.Value = DLookup("[RF Camber]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtRFToe.Value = DLookup("[RF Toe]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtRFShock.Value = DLookup("[RF Shock]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtRFPressureCold.Value = DLookup("[RF Pressure Cold]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtRFPressureHot.Value = DLookup("[RF Pressure Hot]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    'RR
    Form_Setup.txtRRWeight.Value = DLookup("[RR Weight]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtRRRideHeight.Value = DLookup("[RR Ride Height]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtRRCamber.Value = DLookup("[RR Camber]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtRRToe.Value = DLookup("[RR Toe]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtRRShock.Value = DLookup("[RR Shock]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtRRPressureCold.Value = DLookup("[RR Pressure Cold]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    Form_Setup.txtRRPressureHot.Value = DLookup("[RR Pressure Hot]", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
    'Notes
    Form_Setup.txtNote.Value = DLookup("Notes", "tblCarSetup", "ID=" & Form_SubCarSetup.ID.Value)
   On Error GoTo 0
    
End Sub
