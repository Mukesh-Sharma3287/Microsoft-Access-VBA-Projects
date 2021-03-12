VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SubRaceSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
    On Error Resume Next
    Form_Race_Session.lblID.Caption = Form_SubRaceSession.ID.Value
    Form_Race_Session.txtDate.Value = Form_SubRaceSession.RaceDate.Value
    Form_Race_Session.ddnTrack.Value = Form_SubRaceSession.Track.Value
    Form_Race_Session.ddnEvent.Value = Form_SubRaceSession.Event.Value
    Form_Race_Session.txtCar.Value = Form_SubRaceSession.Car.Value
    Form_Race_Session.txtBaseCar.Value = Form_SubRaceSession.Base_Car_Setup.Value
    Form_Race_Session.ddnCrewChief.Value = Form_SubRaceSession.Crew_Chief.Value
    Form_Race_Session.ddnMember1.Value = Form_SubRaceSession.Member1.Value
    Form_Race_Session.txtMember2.Value = Form_SubRaceSession.Member2.Value
    Form_Race_Session.txtMember3.Value = Form_SubRaceSession.Member3.Value
    Form_Race_Session.txtMember4.Value = Form_SubRaceSession.Member4.Value
    Form_Race_Session.txtDuration.Value = Form_SubRaceSession.Duration.Value
    Form_Race_Session.ddnType.Value = Form_SubRaceSession.Type.Value
    Form_Race_Session.txtTime.Value = Form_SubRaceSession.Time.Value
    Form_Race_Session.txtAirTemp.Value = Form_SubRaceSession.Air_Temp.Value
    Form_Race_Session.txtTrackTemp.Value = Form_SubRaceSession.Track_Temp.Value
    Form_Race_Session.txtAirDensity.Value = Form_SubRaceSession.Air_Density.Value
    Form_Race_Session.txtHumidity.Value = Form_SubRaceSession.Humidity.Value
    Form_Race_Session.txtBarPressure.Value = Form_SubRaceSession.Bar_Pressure.Value
    Form_Race_Session.ddnWeather.Value = Form_SubRaceSession.Weather.Value
    Form_Race_Session.txtTireIDLF.Value = Form_SubRaceSession.Tire_ID_LF.Value
    Form_Race_Session.txtLFPressOut.Value = Form_SubRaceSession.LF_Press_Out.Value
    Form_Race_Session.txtLFPressIn.Value = Form_SubRaceSession.LF_Press_In.Value
    Form_Race_Session.txtLFOutTemp.Value = Form_SubRaceSession.LF_Outside.Value
    Form_Race_Session.txtLFMidTemp.Value = Form_SubRaceSession.LF_Middle.Value
    Form_Race_Session.txtLFInsideTemp.Value = Form_SubRaceSession.LF_Inside.Value
    Form_Race_Session.txtLFAvgTemp.Value = Form_SubRaceSession.LF_Avg.Value
    Form_Race_Session.txtTireIDRF.Value = Form_SubRaceSession.Tire_ID_RF.Value
    Form_Race_Session.txtRFPressOut.Value = Form_SubRaceSession.RF_Press_Out.Value
    Form_Race_Session.txtRFPressIn.Value = Form_SubRaceSession.RF_Press_In.Value
    Form_Race_SessiontxtRFOutside.Value = Form_SubRaceSession.RF_Outside.Value
    Form_Race_Session.txtRFMid.Value = Form_SubRaceSession.RF_Middle.Value
    Form_Race_Session.txtRFInside.Value = Form_SubRaceSession.RF_Inside.Value
    Form_Race_Session.txtRFAvg.Value = Form_SubRaceSession.RF_Avg.Value
    Form_Race_Session.txtTireIDLR.Value = Form_SubRaceSession.Tire_ID_LR.Value
    Form_Race_Session.txtLRPressOut.Value = Form_SubRaceSession.LR_Press_Out.Value
    Form_Race_Session.txtLRPressIn.Value = Form_SubRaceSession.LR_Press_In.Value
    Form_Race_Session.txtLROutside.Value = Form_SubRaceSession.LR_Outside.Value
    Form_Race_Session.txtLRMid.Value = Form_SubRaceSession.LR_Middle.Value
    Form_Race_Session.txtLRInside.Value = Form_SubRaceSession.LR_Inside.Value
    Form_Race_Session.txtLRAvg.Value = Form_SubRaceSession.LR_Avg.Value
    Form_Race_Session.txtTireIDRR.Value = Form_SubRaceSession.Tire_ID_RR.Value
    Form_Race_Session.txtRRPressOut.Value = Form_SubRaceSession.RR_Press_Out.Value
    Form_Race_Session.txtRRPressIn.Value = Form_SubRaceSession.RR_Press_In.Value
    Form_Race_Session.txtRROutside.Value = Form_SubRaceSession.RR_Outside.Value
    Form_Race_Session.txtRRMid.Value = Form_SubRaceSession.RR_Middle.Value
    Form_Race_Session.txtRFInside.Value = Form_SubRaceSession.RF_Inside.Value
    Form_Race_Session.txtRFAvg.Value = Form_SubRaceSession.RF_Avg.Value
    Form_Race_Session.txtLFShock.Value = Form_SubRaceSession.LF_Shock.Value
    Form_Race_Session.txtLRShock.Value = Form_SubRaceSession.LR_Shock.Value
    Form_Race_Session.txtRFShock.Value = Form_SubRaceSession.RF_Shock.Value
    Form_Race_Session.txtRRShock.Value = Form_SubRaceSession.RR_Shock.Value
    Form_Race_Session.txtFuelIn.Value = Form_SubRaceSession.Fuel_In.Value
    Form_Race_Session.txtFuelOut.Value = Form_SubRaceSession.Fuel_Out.Value
   
    On Error GoTo 0
    
End Sub


