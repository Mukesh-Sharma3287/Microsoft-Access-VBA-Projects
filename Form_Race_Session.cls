VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Race_Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdAddNew_Click()
    
    'Call Function to Submit Details
    Call ModSave.fnSubmitRaceSession("Save")
    
End Sub

Private Sub cmdDelete_Click()
    
     'Call Function to Delete record
    Call ModSave.DeleteRaceData
    
End Sub

Private Sub cmdReset_Click()
    
    'Call Function to Delete record
    Call ModSave.ResetData
    
End Sub

Private Sub cmdTime_Click()
    Form_Race_Session.txtTime.Value = Format(Now(), "h:mm:ss AM/PM")
End Sub

Private Sub cmdUpdate_Click()
    
    'Call Function to Update Details
    Call ModSave.fnSubmitRaceSession("Update")
    
End Sub

Private Sub ddnCrewChief_AfterUpdate()
    If ModSave.fnCheckUpdate("Crew Chief", lblID.Caption, Form_Race_Session.ddnCrewChief.Value) = True Then
        Form_Race_Session.ddnCrewChief.BackColor = vbGreen
    End If
End Sub

Private Sub ddnEvent_AfterUpdate()
   If ModSave.fnCheckUpdate("Event", lblID.Caption, Form_Race_Session.ddnEvent.Value) = True Then
        Form_Race_Session.ddnEvent.BackColor = vbGreen
    End If
End Sub

Private Sub ddnMember1_AfterUpdate()
     If ModSave.fnCheckUpdate("Member1", lblID.Caption, Form_Race_Session.ddnMember1.Value) = True Then
        Form_Race_Session.ddnMember1.BackColor = vbGreen
    End If
End Sub

Private Sub ddnMember2_AfterUpdate()
     If ModSave.fnCheckUpdate("Member2", lblID.Caption, Form_Race_Session.ddnMember2.Value) = True Then
        Form_Race_Session.ddnMember2.BackColor = vbGreen
    End If
End Sub

Private Sub ddnMember3_AfterUpdate()
     If ModSave.fnCheckUpdate("Member3", lblID.Caption, Form_Race_Session.ddnMember3.Value) = True Then
        Form_Race_Session.ddnMember3.BackColor = vbGreen
    End If
End Sub

Private Sub ddnMember4_AfterUpdate()
     If ModSave.fnCheckUpdate("Member4", lblID.Caption, Form_Race_Session.ddnMember4.Value) = True Then
        Form_Race_Session.ddnMember4.BackColor = vbGreen
    End If
End Sub

Private Sub ddnType_AfterUpdate()
    If ModSave.fnCheckUpdate("Type", lblID.Caption, Form_Race_Session.ddnType.Value) = True Then
        Form_Race_Session.ddnType.BackColor = vbGreen
    End If
End Sub

Private Sub ddnWeather_AfterUpdate()
     If ModSave.fnCheckUpdate("Weather", lblID.Caption, Form_Race_Session.ddnWeather.Value) = True Then
        Form_Race_Session.ddnWeather.BackColor = vbGreen
    End If
End Sub

Private Sub Form_Load()
    Call fnResetColor
    Form_Race_Session.txtTime.Value = Format(Now(), "h:mm:ss AM/PM")
    
End Sub

Private Sub txtAirDensity_AfterUpdate()
    If ModSave.fnCheckUpdate("Air Density", lblID.Caption, Form_Race_Session.txtAirDensity.Value) = True Then
        Form_Race_Session.txtAirDensity.BackColor = vbGreen
    End If
End Sub

Private Sub txtAirTemp_AfterUpdate()
    If ModSave.fnCheckUpdate("Air Temp", lblID.Caption, Form_Race_Session.txtAirTemp.Value) = True Then
        Form_Race_Session.txtAirTemp.BackColor = vbGreen
    End If
End Sub

Private Sub txtBarPressure_AfterUpdate()
    If ModSave.fnCheckUpdate("Bar Pressure", lblID.Caption, Form_Race_Session.txtBarPressure.Value) = True Then
        Form_Race_Session.txtBarPressure.BackColor = vbGreen
    End If
End Sub

Private Sub txtBaseCar_AfterUpdate()
     If ModSave.fnCheckUpdate("Base Car Setup", lblID.Caption, Form_Race_Session.txtBaseCar.Value) = True Then
        Form_Race_Session.txtBaseCar.BackColor = vbGreen
    End If
End Sub

Private Sub txtCar_AfterUpdate()
    If ModSave.fnCheckUpdate("Car", lblID.Caption, Form_Race_Session.txtCar.Value) = True Then
        Form_Race_Session.txtCar.BackColor = vbGreen
    End If
End Sub

Private Sub txtDate_AfterUpdate()
   
   If ModSave.fnCheckUpdate("RaceDate", lblID.Caption, Form_Race_Session.txtDate.Value) = True Then
        Form_Race_Session.txtDate.BackColor = vbGreen
    End If
    
End Sub

Private Sub ddnTrack_AfterUpdate()
      If ModSave.fnCheckUpdate("Track", lblID.Caption, Form_Race_Session.ddnTrack.Value) = True Then
        Form_Race_Session.ddnTrack.BackColor = vbGreen
    End If
End Sub


Private Sub txtDuration_AfterUpdate()
     If ModSave.fnCheckUpdate("Duration", lblID.Caption, Form_Race_Session.txtDuration.Value) = True Then
        Form_Race_Session.txtDuration.BackColor = vbGreen
    End If
End Sub

Private Sub txtFuelIn_AfterUpdate()
    If ModSave.fnCheckUpdate("Fuel In", lblID.Caption, Form_Race_Session.txtFuelIn.Value) = True Then
        Form_Race_Session.txtFuelIn.BackColor = vbGreen
    End If
End Sub

Private Sub txtFuelOut_AfterUpdate()
    If ModSave.fnCheckUpdate("Fuel Out", lblID.Caption, Form_Race_Session.txtFuelOut.Value) = True Then
        Form_Race_Session.txtFuelOut.BackColor = vbGreen
    End If
End Sub

Private Sub txtHumidity_AfterUpdate()
     If ModSave.fnCheckUpdate("Humidity", lblID.Caption, Form_Race_Session.txtHumidity.Value) = True Then
        Form_Race_Session.txtHumidity.BackColor = vbGreen
    End If
End Sub
'**********************LF****************************************
Private Sub txtLFAvgTemp_AfterUpdate()
     If ModSave.fnCheckUpdate("LF Avg", lblID.Caption, Form_Race_Session.txtLFAvgTemp.Value) = True Then
        Form_Race_Session.txtLFAvgTemp.BackColor = vbGreen
    End If
        
End Sub

Private Sub txtLFInsideTemp_AfterUpdate()
     If ModSave.fnCheckUpdate("LF Inside", lblID.Caption, Form_Race_Session.txtLFInsideTemp.Value) = True Then
        Form_Race_Session.txtLFInsideTemp.BackColor = vbGreen
    End If
    Form_Race_Session.txtLFAvgTemp.Value = Round((Int(Form_Race_Session.txtLFOutTemp.Value) + Int(Form_Race_Session.txtLFMidTemp.Value) + Int(Form_Race_Session.txtLFInsideTemp.Value)) / 3, 2)
End Sub

Private Sub txtLFMidTemp_AfterUpdate()
     If ModSave.fnCheckUpdate("LF Middle", lblID.Caption, Form_Race_Session.txtLFMidTemp.Value) = True Then
        Form_Race_Session.txtLFMidTemp.BackColor = vbGreen
    End If
    
    Form_Race_Session.txtLFAvgTemp.Value = Round((Int(Form_Race_Session.txtLFOutTemp.Value) + Int(Form_Race_Session.txtLFMidTemp.Value) + Int(Form_Race_Session.txtLFInsideTemp.Value)) / 3, 2)
End Sub

Private Sub txtLFOutTemp_AfterUpdate()
     If ModSave.fnCheckUpdate("LF Outside", lblID.Caption, Form_Race_Session.txtLFOutTemp.Value) = True Then
        Form_Race_Session.txtLFOutTemp.BackColor = vbGreen
    End If
    
    Form_Race_Session.txtLFAvgTemp.Value = Round((Int(Form_Race_Session.txtLFOutTemp.Value) + Int(Form_Race_Session.txtLFMidTemp.Value) + Int(Form_Race_Session.txtLFInsideTemp.Value)) / 3, 2)
End Sub

Private Sub txtLFPressIn_AfterUpdate()
     If ModSave.fnCheckUpdate("LF Press In", lblID.Caption, Form_Race_Session.txtLFPressIn.Value) = True Then
        Form_Race_Session.txtLFPressIn.BackColor = vbGreen
    End If
End Sub

Private Sub txtLFPressOut_AfterUpdate()
     If ModSave.fnCheckUpdate("LF Press Out", lblID.Caption, Form_Race_Session.txtLFPressOut.Value) = True Then
        Form_Race_Session.txtLFPressOut.BackColor = vbGreen
    End If
End Sub

'*************Shock*************************
Private Sub txtLFShock_AfterUpdate()
    If ModSave.fnCheckUpdate("LF Shock", lblID.Caption, Form_Race_Session.txtLFShock.Value) = True Then
        Form_Race_Session.txtLFShock.BackColor = vbGreen
    End If
End Sub



Private Sub txtLRShock_AfterUpdate()
    If ModSave.fnCheckUpdate("LR Shock", lblID.Caption, Form_Race_Session.txtLRShock.Value) = True Then
        Form_Race_Session.txtLRShock.BackColor = vbGreen
    End If
    
    Form_Race_Session.txtLRAvgTemp.Value = Round((Form_Race_Session.txtLROutTemp.Value + Form_Race_Session.txtLRMidTemp.Value + Form_Race_Session.txtLRInsideTemp.Value) / 3, 2)
End Sub



Private Sub txtMember4_AfterUpdate()
     If ModSave.fnCheckUpdate("Member4", lblID.Caption, Form_Race_Session.txtMember4.Value) = True Then
        Form_Race_Session.txtMember4.BackColor = vbGreen
    End If
End Sub

Private Sub txtMember3_AfterUpdate()
     If ModSave.fnCheckUpdate("Member3", lblID.Caption, Form_Race_Session.txtMember3.Value) = True Then
        Form_Race_Session.txtMember3.BackColor = vbGreen
    End If
End Sub

Private Sub txtMember2_AfterUpdate()
     If ModSave.fnCheckUpdate("Member2", lblID.Caption, Form_Race_Session.txtMember2.Value) = True Then
        Form_Race_Session.txtMember2.BackColor = vbGreen
    End If
End Sub

Private Sub txtRFShock_AfterUpdate()
    If ModSave.fnCheckUpdate("RF Shock", lblID.Caption, Form_Race_Session.txtRFShock.Value) = True Then
        Form_Race_Session.txtRFShock.BackColor = vbGreen
    End If
End Sub

Private Sub txtRRShock_AfterUpdate()
    If ModSave.fnCheckUpdate("RR Shock", lblID.Caption, Form_Race_Session.txtRRShock.Value) = True Then
        Form_Race_Session.txtRRShock.BackColor = vbGreen
    End If
End Sub

Private Sub txtSessionNotes_AfterUpdate()
    If ModSave.fnCheckUpdate("Session Notes", lblID.Caption, Form_Race_Session.txtSessionNotes.Value) = True Then
        Form_Race_Session.txtSessionNotes.BackColor = vbGreen
    End If
End Sub

Private Sub txtTireIDLF_AfterUpdate()
     If ModSave.fnCheckUpdate("Tire ID LF", lblID.Caption, Form_Race_Session.txtTireIDLF.Value) = True Then
        Form_Race_Session.txtTireIDLF.BackColor = vbGreen
    End If
End Sub


'**********************RF****************************************
Private Sub txtRFAvg_AfterUpdate()
     If ModSave.fnCheckUpdate("RF Avg", lblID.Caption, Form_Race_Session.txtRFAvg.Value) = True Then
        Form_Race_Session.tx tRFAvgTemp.BackColor = vbGreen
    End If
End Sub

Private Sub txtRFInside_AfterUpdate()
     If ModSave.fnCheckUpdate("RF Inside", lblID.Caption, Form_Race_Session.txtRFInside.Value) = True Then
        Form_Race_Session.txtRFInside.BackColor = vbGreen
    End If
    
    Form_Race_Session.txtRFAvg.Value = Round((Int(Form_Race_Session.txtRFOutside.Value) + Int(Form_Race_Session.txtRFMid.Value) + Int(Form_Race_Session.txtRFInside.Value)) / 3, 2)
End Sub

Private Sub txtRFMid_AfterUpdate()
     If ModSave.fnCheckUpdate("RF Middle", lblID.Caption, Form_Race_Session.txtRFMid.Value) = True Then
        Form_Race_Session.txtRFMid.BackColor = vbGreen
    End If
    
    Form_Race_Session.txtRFAvg.Value = Round((Int(Form_Race_Session.txtRFOutside.Value) + Int(Form_Race_Session.txtRFMid.Value) + Int(Form_Race_Session.txtRFInside.Value)) / 3, 2)
End Sub

Private Sub txtRFOutside_AfterUpdate()
     If ModSave.fnCheckUpdate("RF Outside", lblID.Caption, Form_Race_Session.txtRFOutside.Value) = True Then
        Form_Race_Session.txtRFOutside.BackColor = vbGreen
    End If
    Form_Race_Session.txtRFAvg.Value = Round((Int(Form_Race_Session.txtRFOutside.Value) + Int(Form_Race_Session.txtRFMid.Value) + Int(Form_Race_Session.txtRFInside.Value)) / 3, 2)
End Sub

Private Sub txtRFPressIn_AfterUpdate()
     If ModSave.fnCheckUpdate("RF Press In", lblID.Caption, Form_Race_Session.txtRFPressIn.Value) = True Then
        Form_Race_Session.txtRFPressIn.BackColor = vbGreen
    End If
End Sub

Private Sub txtRFPressOut_AfterUpdate()
     If ModSave.fnCheckUpdate("RF Press Out", lblID.Caption, Form_Race_Session.txtRFPressOut.Value) = True Then
        Form_Race_Session.txtRFPressOut.BackColor = vbGreen
    End If
End Sub

Private Sub txtTireIDRF_AfterUpdate()
     If ModSave.fnCheckUpdate("Tire ID RF", lblID.Caption, Form_Race_Session.txtTireIDRF.Value) = True Then
        Form_Race_Session.txtTireIDRF.BackColor = vbGreen
    End If
End Sub



'**********************LR****************************************
Private Sub txtLRAvg_AfterUpdate()
     If ModSave.fnCheckUpdate("LR Avg", lblID.Caption, Form_Race_Session.txtLRAvg.Value) = True Then
        Form_Race_Session.txtLRAvg.BackColor = vbGreen
    End If
End Sub

Private Sub txtLRInside_AfterUpdate()
     If ModSave.fnCheckUpdate("LR Inside", lblID.Caption, Form_Race_Session.txtLRInside.Value) = True Then
        Form_Race_Session.txtLRInside.BackColor = vbGreen
    End If
    
    Form_Race_Session.txtLRAvg.Value = Round((Int(Form_Race_Session.txtLROutside.Value) + Int(Form_Race_Session.txtLRMid.Value) + Int(Form_Race_Session.txtLRInside.Value)) / 3, 2)
End Sub

Private Sub txtLRMid_AfterUpdate()
     If ModSave.fnCheckUpdate("LR Middle", lblID.Caption, Form_Race_Session.txtLRMid.Value) = True Then
        Form_Race_Session.txtLRMid.BackColor = vbGreen
    End If
    
    Form_Race_Session.txtLRAvg.Value = Round((Int(Form_Race_Session.txtLROutside.Value) + Int(Form_Race_Session.txtLRMid.Value) + Int(Form_Race_Session.txtLRInside.Value)) / 3, 2)
End Sub

Private Sub txtLROutside_AfterUpdate()
     If ModSave.fnCheckUpdate("LR Outside", lblID.Caption, Form_Race_Session.txtLROutside.Value) = True Then
        Form_Race_Session.txtLROutside.BackColor = vbGreen
    End If
    
    Form_Race_Session.txtLRAvg.Value = Round((Int(Form_Race_Session.txtLROutside.Value) + Int(Form_Race_Session.txtLRMid.Value) + Int(Form_Race_Session.txtLRInside.Value)) / 3, 2)
End Sub

Private Sub txtLRPressIn_AfterUpdate()
     If ModSave.fnCheckUpdate("LR Press In", lblID.Caption, Form_Race_Session.txtLRPressIn.Value) = True Then
        Form_Race_Session.txtLRPressIn.BackColor = vbGreen
    End If
End Sub

Private Sub txtLRPressOut_AfterUpdate()
     If ModSave.fnCheckUpdate("LR Press Out", lblID.Caption, Form_Race_Session.txtLRPressOut.Value) = True Then
        Form_Race_Session.txtLRPressOut.BackColor = vbGreen
    End If
End Sub

Private Sub txtTireIDLR_AfterUpdate()
     If ModSave.fnCheckUpdate("Tire ID LR", lblID.Caption, Form_Race_Session.txtTireIDLR.Value) = True Then
        Form_Race_Session.txtTireIDLR.BackColor = vbGreen
    End If
End Sub



'**********************RR****************************************
Private Sub txtRRAvg_AfterUpdate()
     If ModSave.fnCheckUpdate("RR Avg", lblID.Caption, Form_Race_Session.txtRRAvg.Value) = True Then
        Form_Race_Session.txtRRAvg.BackColor = vbGreen
    End If
End Sub

Private Sub txtRRInside_AfterUpdate()
     If ModSave.fnCheckUpdate("RR Inside", lblID.Caption, Form_Race_Session.txtRRInside.Value) = True Then
        Form_Race_Session.txtRRInside.BackColor = vbGreen
    End If
    
    Form_Race_Session.txtRRAvg.Value = Round((Int(Form_Race_Session.txtRROutside.Value) + Int(Form_Race_Session.txtRRMid.Value) + Int(Form_Race_Session.txtRRInside.Value)) / 3, 2)
End Sub

Private Sub txtRRMid_AfterUpdate()
     If ModSave.fnCheckUpdate("RR Middle", lblID.Caption, Form_Race_Session.txtRRMid.Value) = True Then
        Form_Race_Session.txtRRMid.BackColor = vbGreen
    End If
    
    Form_Race_Session.txtRRAvg.Value = Round((Int(Form_Race_Session.txtRROutside.Value) + Int(Form_Race_Session.txtRRMid.Value) + Int(Form_Race_Session.txtRRInside.Value)) / 3, 2)
End Sub

Private Sub txtRROutside_AfterUpdate()
     If ModSave.fnCheckUpdate("RR Outside", lblID.Caption, Form_Race_Session.txtRROutside.Value) = True Then
        Form_Race_Session.txtRROutside.BackColor = vbGreen
    End If
    
    Form_Race_Session.txtRRAvg.Value = Round((Int(Form_Race_Session.txtRROutside.Value) + Int(Form_Race_Session.txtRRMid.Value) + Int(Form_Race_Session.txtRRInside.Value)) / 3, 2)
End Sub

Private Sub txtRRPressIn_AfterUpdate()
     If ModSave.fnCheckUpdate("RR Press In", lblID.Caption, Form_Race_Session.txtRRPressIn.Value) = True Then
        Form_Race_Session.txtRRPressIn.BackColor = vbGreen
    End If
End Sub

Private Sub txtRRPressOut_AfterUpdate()
     If ModSave.fnCheckUpdate("RR Press Out", lblID.Caption, Form_Race_Session.txtRRPressOut.Value) = True Then
        Form_Race_Session.txtRRPressOut.BackColor = vbGreen
    End If
End Sub

Private Sub txtTireIDRR_AfterUpdate()
     If ModSave.fnCheckUpdate("Tire ID RR", lblID.Caption, Form_Race_Session.txtTireIDRR.Value) = True Then
        Form_Race_Session.txtTireIDRR.BackColor = vbGreen
    End If
End Sub



Private Sub txtTrackTemp_AfterUpdate()
    If ModSave.fnCheckUpdate("Track Temp", lblID.Caption, Form_Race_Session.txtTrackTemp.Value) = True Then
        Form_Race_Session.txtTrackTemp.BackColor = vbGreen
    End If
End Sub
Private Sub txtTime_AfterUpdate()
    If ModSave.fnCheckUpdate("Time", lblID.Caption, Form_Race_Session.txtTime.Value) = True Then
        Form_Race_Session.txtTime.BackColor = vbGreen
    End If
End Sub
