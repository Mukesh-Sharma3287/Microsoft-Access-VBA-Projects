Attribute VB_Name = "ModSave"
Option Compare Database
Option Explicit

''***********************************************************************************************************
'' Form Name        : ModFunction
'' Description      : This Module contain the Various function to use the entire tool
''
''
''    Date                  Developer               Created             Remarks
''  15-Nov-2020            Mukesh Sharma             First Version
''***********************************************************************************************************

''***********************************************************************************************************
'' Procedure Name       : fnSubmitTransactions
'' Description          : To submit the transaction details
'' Arguments            : NA
'' Returns              : NA
''***********************************************************************************************************

Public Function fnSubmitRaceSession(strStatus As String)
    
    Dim rsRecordset As Recordset
    Dim strQry As String
             
     'Validation
     
     If Form_Race_Session.txtTireIDLF.Value = Form_Race_Session.txtTireIDRF.Value Or Form_Race_Session.txtTireIDRF.Value = Form_Race_Session.txtTireIDLR.Value Or Form_Race_Session.txtTireIDLR.Value = Form_Race_Session.txtTireIDRR.Value Or Form_Race_Session.txtTireIDRR.Value = Form_Race_Session.txtTireIDLF.Value Then
        MsgBox "Please select the different Tire ID", vbInformation
        Exit Function
     ElseIf Form_Race_Session.txtDuration.Value = "" Or IsNull(Form_Race_Session.txtDuration.Value) = True Then
        MsgBox "Please select the Duration", vbInformation
        Exit Function
      ElseIf Form_Race_Session.ddnTrack.Value = "" Or IsNull(Form_Race_Session.ddnTrack.Value) = True Then
        MsgBox "Please select the Track", vbInformation
        Exit Function
        
     ElseIf Form_Race_Session.ddnEvent.Value = "" Or IsNull(Form_Race_Session.ddnEvent.Value) = True Then
        MsgBox "Please select the Event", vbInformation
        Exit Function
        
     ElseIf Form_Race_Session.txtBaseCar.Value = "" Or IsNull(Form_Race_Session.txtBaseCar.Value) = True Then
        MsgBox "Please select the Base Car", vbInformation
        Exit Function
      ElseIf Form_Race_Session.txtTime.Value = "" Or IsNull(Form_Race_Session.txtTime.Value) = True Then
        MsgBox "Please Enter the Time", vbInformation
        Exit Function
        
     End If
   
    strQry = "Select * from tblRaceSession where ID=" & IIf(Form_Race_Session.lblID.Caption = "", 0, Form_Race_Session.lblID.Caption)
    
    Set rsRecordset = CurrentDb.OpenRecordset(strQry)
    
    With rsRecordset
    
        If strStatus = "Save" Then
             .AddNew
        Else
              .Edit
        End If
        
        ![RaceDate] = Form_Race_Session.txtDate.Value
        ![Track] = Form_Race_Session.ddnTrack.Value
        ![Event] = Form_Race_Session.ddnEvent.Value
        ![Car] = Form_Race_Session.txtCar.Value
        ![Base Car Setup] = Form_Race_Session.txtBaseCar.Value
        ![Crew Chief] = Form_Race_Session.ddnCrewChief.Value
        ![Member1] = Form_Race_Session.ddnMember1.Value
        ![Member2] = Form_Race_Session.txtMember2.Value
        ![Member3] = Form_Race_Session.txtMember3.Value
        ![Member4] = Form_Race_Session.txtMember4.Value
        ![Duration] = Form_Race_Session.txtDuration.Value
        ![Type] = Form_Race_Session.ddnType.Value
        ![Time] = Form_Race_Session.txtTime.Value
        ![Air Temp] = Form_Race_Session.txtAirTemp.Value
        ![Track Temp] = Form_Race_Session.txtTrackTemp.Value
        ![Air Density] = Form_Race_Session.txtAirDensity.Value
        ![Humidity] = Form_Race_Session.txtHumidity.Value
        ![Bar Pressure] = Form_Race_Session.txtBarPressure.Value
        ![Weather] = Form_Race_Session.ddnWeather.Value
        ![Tire ID LF] = Form_Race_Session.txtTireIDLF.Value
        
        'LF
         Call fnInsertTireID(Form_Race_Session.lblID.Caption, Form_Race_Session.txtTireIDLF.Value, Form_Race_Session.txtDuration.Value, "LF", Form_Race_Session.ddnTrack.Value, Form_Race_Session.txtDate.Value, Form_Race_Session.ddnEvent.Value, Form_Race_Session.txtBaseCar.Value, Form_Race_Session.txtTime.Value)
         
        ![LF Press Out] = Form_Race_Session.txtLFPressOut.Value
        ![LF Press In] = Form_Race_Session.txtLFPressIn.Value
        ![LF Outside] = Form_Race_Session.txtLFOutTemp.Value
        ![LF Middle] = Form_Race_Session.txtLFMidTemp.Value
        ![LF Inside] = Form_Race_Session.txtLFInsideTemp.Value
        ![LF Avg] = (Int(Form_Race_Session.txtLFOutTemp.Value) + Int(Form_Race_Session.txtLFMidTemp.Value) + Int(Form_Race_Session.txtLFInsideTemp.Value)) / 3
        
        
        ![Tire ID RF] = Form_Race_Session.txtTireIDRF.Value
        
        'RF
        Call fnInsertTireID(Form_Race_Session.lblID.Caption, Form_Race_Session.txtTireIDRF.Value, Form_Race_Session.txtDuration.Value, "RF", Form_Race_Session.ddnTrack.Value, Form_Race_Session.txtDate.Value, Form_Race_Session.ddnEvent.Value, Form_Race_Session.txtBaseCar.Value, Form_Race_Session.txtTime.Value)
        
        ![RF Press Out] = Form_Race_Session.txtRFPressOut.Value
        ![RF Press In] = Form_Race_Session.txtRFPressIn.Value
        ![RF Outside] = Form_Race_Session.txtRFOutside.Value
        ![RF Middle] = Form_Race_Session.txtRFMid.Value
        ![RF Inside] = Form_Race_Session.txtRFInside.Value
        ![RF Avg] = (Int(Form_Race_Session.txtRFOutside.Value) + Int(Form_Race_Session.txtRFMid.Value) + Int(Form_Race_Session.txtRFInside.Value)) / 3
        
        ![Tire ID LR] = Form_Race_Session.txtTireIDLR.Value
        
        'LR
        Call fnInsertTireID(Form_Race_Session.lblID.Caption, Form_Race_Session.txtTireIDLR.Value, Form_Race_Session.txtDuration.Value, "LR", Form_Race_Session.ddnTrack.Value, Form_Race_Session.txtDate.Value, Form_Race_Session.ddnEvent.Value, Form_Race_Session.txtBaseCar.Value, Form_Race_Session.txtTime.Value)
        
        ![LR Press Out] = Form_Race_Session.txtLRPressOut.Value
        ![LR Press In] = Form_Race_Session.txtLRPressIn.Value
        ![LR Outside] = Form_Race_Session.txtLROutside.Value
        ![LR Middle] = Form_Race_Session.txtLRMid.Value
        ![LR Inside] = Form_Race_Session.txtLRInside.Value
        ![LR Avg] = (Int(Form_Race_Session.txtLROutside.Value) + Int(Form_Race_Session.txtLRMid.Value) + Int(Form_Race_Session.txtLRInside.Value)) / 3
        ![Tire ID RR] = Form_Race_Session.txtTireIDRR.Value
        
         'RR
        Call fnInsertTireID(Form_Race_Session.lblID.Caption, Form_Race_Session.txtTireIDRR.Value, Form_Race_Session.txtDuration.Value, "RR", Form_Race_Session.ddnTrack.Value, Form_Race_Session.txtDate.Value, Form_Race_Session.ddnEvent.Value, Form_Race_Session.txtBaseCar.Value, Form_Race_Session.txtTime.Value)
        
        ![RR Press Out] = Form_Race_Session.txtRRPressOut.Value
        ![RR Press In] = Form_Race_Session.txtRRPressIn.Value
        ![RR Outside] = Form_Race_Session.txtRROutside.Value
        ![RR Middle] = Form_Race_Session.txtRRMid.Value
        ![RR Inside] = Form_Race_Session.txtRRInside.Value
        ![RR Avg] = (Int(Form_Race_Session.txtRFInside.Value) + Int(Form_Race_Session.txtRRMid.Value) + Int(Form_Race_Session.txtRRMid.Value)) / 3
        ![LF Shock] = Form_Race_Session.txtLFShock.Value
        ![LR Shock] = Form_Race_Session.txtLRShock.Value
        ![RF Shock] = Form_Race_Session.txtRFShock.Value
        ![RR Shock] = Form_Race_Session.txtRRShock.Value
        ![Fuel Out] = Form_Race_Session.txtFuelIn.Value
        ![Fuel In] = Form_Race_Session.txtFuelOut.Value
        .Update
    End With
  
     On Error GoTo 0
     
    If Form_Race_Session.cmdAddNew.Caption = "Save" Then
        MsgBox "Record has been saved successfully!!!", vbInformation
    Else
         MsgBox "Record has been updated successfully!!!", vbInformation
    End If

    'Refresh
    Form_SubRaceSession.Requery
     
    'Reset data
    Call ModSave.ResetData
    Call fnResetColor
  
End Function

Public Function fnInsertTireID(lID As Long, strTireID As String, lDuration As Long, strType As String, strTrack As String, strRaceDate As String, strEvent As String, strBaseCarSet As String, strTime As String)
    
    On Error GoTo ErrInsert
    Dim strQry As String
    strQry = "INSERT INTO tblTireIDCombined(ID,[Tire ID],[Duration],[Type],[Track],[RaceDate],[Event],[Base Car Set],[Time]) Values(" & lID & ",'" & strTireID & "'," & lDuration & ",'" & strType & "','" & strTrack & "','" & strRaceDate & "','" & strEvent & "','" & strBaseCarSet & "','" & strTime & "')"
    CurrentDb.Execute strQry
    On Error GoTo 0
    
    Exit Function
    
ErrInsert:
    MsgBox Err.Description, vbCritical
    
End Function
Sub ResetData()
    
    Form_Race_Session.lblID.Caption = 0
    Form_Race_Session.txtDate.Value = Empty
    Form_Race_Session.ddnTrack.Value = Empty
    Form_Race_Session.ddnEvent.Value = Empty
    Form_Race_Session.txtCar.Value = Empty
    Form_Race_Session.txtBaseCar.Value = Empty
    Form_Race_Session.ddnCrewChief.Value = Empty
    Form_Race_Session.ddnMember1.Value = Empty
    Form_Race_Session.txtMember2.Value = Empty
    Form_Race_Session.txtMember3.Value = Empty
    Form_Race_Session.txtMember4.Value = Empty
    Form_Race_Session.txtDuration.Value = Empty
    Form_Race_Session.ddnType.Value = Empty
    Form_Race_Session.txtTime.Value = Empty
    Form_Race_Session.txtAirTemp.Value = Empty
    Form_Race_Session.txtTrackTemp.Value = Empty
    Form_Race_Session.txtAirDensity.Value = Empty
    Form_Race_Session.txtHumidity.Value = Empty
    Form_Race_Session.txtBarPressure.Value = Empty
    Form_Race_Session.ddnWeather.Value = Empty
    Form_Race_Session.txtTireIDLF.Value = Empty
    Form_Race_Session.txtLFPressOut.Value = Empty
    Form_Race_Session.txtLFPressIn.Value = Empty
    Form_Race_Session.txtLFOutTemp.Value = Empty
    Form_Race_Session.txtLFMidTemp.Value = Empty
    Form_Race_Session.txtLFInsideTemp.Value = Empty
    Form_Race_Session.txtLFAvgTemp.Value = Empty
    Form_Race_Session.txtTireIDRF.Value = Empty
    Form_Race_Session.txtRFPressOut.Value = Empty
    Form_Race_Session.txtRFPressIn.Value = Empty
    Form_Race_Session.txtRFOutside.Value = Empty
    Form_Race_Session.txtRFMid.Value = Empty
    Form_Race_Session.txtRFInside.Value = Empty
    Form_Race_Session.txtRFAvg.Value = Empty
    Form_Race_Session.txtTireIDLR.Value = Empty
    Form_Race_Session.txtLRPressOut.Value = Empty
    Form_Race_Session.txtLRPressIn.Value = Empty
    Form_Race_Session.txtLROutside.Value = Empty
    Form_Race_Session.txtLRMid.Value = Empty
    Form_Race_Session.txtLRInside.Value = Empty
    Form_Race_Session.txtLRAvg.Value = Empty
    Form_Race_Session.txtTireIDRR.Value = Empty
    Form_Race_Session.txtRRPressOut.Value = Empty
    Form_Race_Session.txtRRPressIn.Value = Empty
    Form_Race_Session.txtRROutside.Value = Empty
    Form_Race_Session.txtRRMid.Value = Empty
    Form_Race_Session.txtRFInside.Value = Empty
    Form_Race_Session.txtRFAvg.Value = Empty
    Form_Race_Session.txtLFShock.Value = Empty
    Form_Race_Session.txtLRShock.Value = Empty
    Form_Race_Session.txtRFShock.Value = Empty
    Form_Race_Session.txtRRShock.Value = Empty
    Form_Race_Session.txtFuelIn.Value = Empty
    Form_Race_Session.txtFuelOut.Value = Empty
    
    Call fnResetColor
End Sub

Sub DeleteRaceData()
    
    Dim strQry As String
    
    strQry = "Delete * from tblRaceSession where ID=" & Form_Race_Session.lblID.Caption
    
    CurrentDb.Execute strQry
    
    Form_SubRaceSession.Requery
    
    MsgBox "Record has been Delected successfully!!!", vbInformation
    
End Sub

Sub SubmitMaintenance(strStatus As String)
      
    Dim rsRecordset As Recordset
    Dim strQry As String
             
    strQry = "Select * from tblCarMaintenance where Service_Record=" & Form_Car_Maintenence.lblID.Caption
    
    Set rsRecordset = CurrentDb.OpenRecordset(strQry)

    With rsRecordset
    
        If strStatus = "Save" Then
             .AddNew
             ![Service_Record] = Form_Car_Maintenence.txtServiceRecord.Value
        Else
              .Edit
        End If
        
        
        ![Start_Date] = Form_Car_Maintenence.txtDate.Value
        ![Car] = Form_Car_Maintenence.txtCar.Value
        ![Service_Type] = Form_Car_Maintenence.txtServiceType.Value
        ![Service_By] = Form_Car_Maintenence.txtServiceBy.Value
        
        .Update
    End With
    
    If Form_Car_Maintenence.cmdSave.Caption = "Saved" Then
        MsgBox "Record has been saved successfully!!!", vbInformation
    Else
         MsgBox "Record has been updated successfully!!!", vbInformation
    End If

    'Refresh
    Form_subCarMaintenance.Requery
     
    
    Call ResetMaintenance
    
End Sub

Sub ResetMaintenance()
    
    Form_Car_Maintenence.lblID.Caption = 0
    Form_Car_Maintenence.txtServiceRecord.Value = ""
    Form_Car_Maintenence.txtDate.Value = ""
    Form_Car_Maintenence.txtCar.Value = ""
    Form_Car_Maintenence.txtServiceType.Value = ""
    Form_Car_Maintenence.txtServiceBy.Value = ""
    
    Call SetNewId
End Sub

Function SetNewId()
    
    If IsNull(DLast("Service_Record", "tblCarMaintenance")) = True Then
        Form_Car_Maintenence.txtServiceRecord.Value = 1
    Else
        Form_Car_Maintenence.txtServiceRecord.Value = DLast("Service_Record", "tblCarMaintenance") + 1
    End If
    
End Function


Sub DeleteMaintenanceData()
    
    Dim strQry As String
    
    strQry = "Delete * from tblCarMaintenance where Service_Record=" & Form_Car_Maintenence.lblID.Caption
    
    CurrentDb.Execute strQry
    
    Form_subCarMaintenance.Requery
    
    MsgBox "Record has been Delected successfully!!!", vbInformation
    
End Sub

Sub submitCarSetup(strStatus As String)
     
    Dim rsRecordset As Recordset
    Dim strQry As String
             
    strQry = "Select * from tblCarSetup where ID=" & Form_Setup.lblID.Caption
    
    Set rsRecordset = CurrentDb.OpenRecordset(strQry)

    With rsRecordset
    
        If strStatus = "Save" Then
             .AddNew
        Else
              .Edit
        End If
        
        ![Setup Id] = Form_Setup.txtSetupId.Value
        ![Car_Type] = Form_Setup.txtCarType.Value
        ![Track Specific] = Form_Setup.txtTrackSpecific.Value
        ![Track Condition] = Form_Setup.txtTrackCondition.Value
        ![Temprature Range] = Form_Setup.txtTempratureRange.Value
        ![Type] = Form_Setup.txtType.Value
        ![Active] = Form_Setup.txtActive.Value
        ![LF Weight] = Form_Setup.txtLFWeight.Value
        ![LF Ride Height] = Form_Setup.txtLFRideHeight.Value
        ![LF Camber] = Form_Setup.txtLFCamber.Value
        ![LF Toe] = Form_Setup.txtLFToe.Value
        ![LF Shock] = Form_Setup.txtLFShock.Value
        ![LF Pressure Cold] = Form_Setup.txtLFPressureCold.Value
        ![LF Pressure Hot] = Form_Setup.txtLFPressureHot.Value
        ![LR Weight] = Form_Setup.txtLRWeight.Value
        ![LR Ride Height] = Form_Setup.txtLRRideHeight.Value
        ![LR Camber] = Form_Setup.txtLRCamber.Value
        ![LR Toe] = Form_Setup.txtLRToe.Value
        ![LR Shock] = Form_Setup.txtLRShock.Value
        ![LR Pressure Cold] = Form_Setup.txtLRPressureCold.Value
        ![LR Pressure Hot] = Form_Setup.txtLRPressureHot.Value
        ![RF Weight] = Form_Setup.txtRFWeight.Value
        ![RF Ride Height] = Form_Setup.txtRFRideHeight.Value
        ![RF Camber] = Form_Setup.txtRFCamber.Value
        ![RF Toe] = Form_Setup.txtRFToe.Value
        ![RF Shock] = Form_Setup.txtRFShock.Value
        ![RF Pressure Cold] = Form_Setup.txtRFPressureCold.Value
        ![RF Pressure Hot] = Form_Setup.txtRFPressureHot.Value
        ![RR Weight] = Form_Setup.txtRRWeight.Value
        ![RR Ride Height] = Form_Setup.txtRRRideHeight.Value
        ![RR Camber] = Form_Setup.txtRRCamber.Value
        ![RR Toe] = Form_Setup.txtRRToe.Value
        ![RR Shock] = Form_Setup.txtRRShock.Value
        ![RR Pressure Cold] = Form_Setup.txtRRPressureCold.Value
        ![RR Pressure Hot] = Form_Setup.txtRRPressureCold.Value
        ![Notes] = Form_Setup.txtNote.Value
        .Update
    End With
     
    If Form_Setup.cmdSave.Caption = "Save" Then
        MsgBox "Record has been saved successfully!!!", vbInformation
    Else
         MsgBox "Record has been updated successfully!!!", vbInformation
    End If

    'Refresh
    Form_SubCarSetup.Requery
    
    Call ResetSetup

End Sub

Sub ResetSetup()
    
    Form_Setup.lblID.Caption = 0
    Form_Setup.txtSetupId.Value = Empty
    Form_Setup.txtCarType.Value = Empty
    Form_Setup.txtTrackSpecific.Value = Empty
    Form_Setup.txtTrackCondition.Value = Empty
    Form_Setup.txtTempratureRange.Value = Empty
    Form_Setup.txtType.Value = Empty
    Form_Setup.txtLFWeight.Value = Empty
    Form_Setup.txtLFRideHeight.Value = Empty
    Form_Setup.txtLFCamber.Value = Empty
    Form_Setup.txtLFToe.Value = Empty
    Form_Setup.txtLFShock.Value = Empty
    Form_Setup.txtLFPressureCold.Value = Empty
    Form_Setup.txtLFPressureHot.Value = Empty
    Form_Setup.txtLRWeight.Value = Empty
    Form_Setup.txtLRRideHeight.Value = Empty
    Form_Setup.txtLRCamber.Value = Empty
    Form_Setup.txtLRToe.Value = Empty
    Form_Setup.txtLRShock.Value = Empty
    Form_Setup.txtLRPressureCold.Value = Empty
    Form_Setup.txtLRPressureHot.Value = Empty
    Form_Setup.txtRFWeight.Value = Empty
    Form_Setup.txtRFRideHeight.Value = Empty
    Form_Setup.txtRFCamber.Value = Empty
    Form_Setup.txtRFToe.Value = Empty
    Form_Setup.txtRFShock.Value = Empty
    Form_Setup.txtRFPressureCold.Value = Empty
    Form_Setup.txtRFPressureHot.Value = Empty
    Form_Setup.txtRRWeight.Value = Empty
    Form_Setup.txtRRRideHeight.Value = Empty
    Form_Setup.txtRRCamber.Value = Empty
    Form_Setup.txtRRToe.Value = Empty
    Form_Setup.txtRRShock.Value = Empty
    Form_Setup.txtRRPressureCold.Value = Empty
    Form_Setup.txtRRPressureHot.Value = Empty
    Form_Setup.txtNote.Value = Empty
    
End Sub

Sub DeleteSetupData()
    
    Dim strQry As String
    
    strQry = "Delete * from tblCarSetup where ID=" & Form_Setup.lblID.Caption
    
    CurrentDb.Execute strQry
    
    Form_SubCarSetup.Requery
    
    MsgBox "Record has been Delected successfully!!!", vbInformation
    
End Sub

Public Function fnTireInventorySave(strStatus As String)
    
    Dim rsRecordset As Recordset
    Dim strQry As String
             
    strQry = "Select * from tblTireInventory where ID=" & Form_Tire_Inventory.lblID.Caption
    
    Set rsRecordset = CurrentDb.OpenRecordset(strQry)

    With rsRecordset
    
        If strStatus = "Save" Then
             .AddNew
        Else
              .Edit
        End If
        ![Tire Id] = Form_Tire_Inventory.txtTireID.Value
        ![Bar Code] = Form_Tire_Inventory.txtBarCode.Value
        ![Set] = Form_Tire_Inventory.txtSet.Value
        ![Manufacturer] = Form_Tire_Inventory.txtManufacturers.Value
        ![Type] = Form_Tire_Inventory.txtType.Value
        ![Designated Car] = Form_Tire_Inventory.txtDesignatedCar.Value
        ![Model Number] = Form_Tire_Inventory.txtModelNumber.Value
        ![Description] = Form_Tire_Inventory.txtDescription.Value
        ![Date Acquired] = Form_Tire_Inventory.txtDataAcquired.Value
        ![Date Mounted] = Form_Tire_Inventory.txtDateMounted.Value
        ![Vendor] = Form_Tire_Inventory.txtVendor.Value
        ![Purchase Invoice] = Form_Tire_Inventory.txtPurchaseInvoice.Value
        ![Cost] = Form_Tire_Inventory.txtCost.Value
        ![Condition Purchased] = Form_Tire_Inventory.txtCondPurchased.Value
        ![Current Condition] = Form_Tire_Inventory.txtCurrentCond.Value
        ![Mounted Rim] = Form_Tire_Inventory.txtMountedRim.Value
        ![Location] = Form_Tire_Inventory.ddnLocation.Value
        ![Active] = Form_Tire_Inventory.ddnActiv.Value
        ![Total Heat Cycle] = Form_Tire_Inventory.txtTotalHeatCycle.Value
        ![Total Time On Tire] = Form_Tire_Inventory.txtTotalTimeOnTire.Value
        ![Note] = Form_Tire_Inventory.txtNote.Value
        .Update
    End With
    
     If Form_Tire_Inventory.cmdSave.Caption = "Save" Then
        MsgBox "Record has been saved successfully!!!", vbInformation
    Else
         MsgBox "Record has been updated successfully!!!", vbInformation
    End If

    'Refresh
    Form_SubTireInventory.Requery
    
    Call ResetTireInventory

End Function

Sub ResetTireInventory()
    
    Form_Tire_Inventory.lblID.Caption = 0
    Form_Tire_Inventory.txtTireID.Value = Empty
    Form_Tire_Inventory.txtBarCode.Value = Empty
    Form_Tire_Inventory.txtSet.Value = Empty
    Form_Tire_Inventory.txtManufacturers.Value = Empty
    Form_Tire_Inventory.txtType.Value = Empty
    Form_Tire_Inventory.txtDesignatedCar.Value = Empty
    Form_Tire_Inventory.txtModelNumber.Value = Empty
    Form_Tire_Inventory.txtDescription.Value = Empty
    Form_Tire_Inventory.txtDataAcquired.Value = Empty
    Form_Tire_Inventory.txtDateMounted.Value = Empty
    Form_Tire_Inventory.txtVendor.Value = Empty
    Form_Tire_Inventory.txtPurchaseInvoice.Value = Empty
    Form_Tire_Inventory.txtCost.Value = Empty
    Form_Tire_Inventory.txtCondPurchased.Value = Empty
    Form_Tire_Inventory.txtCurrentCond.Value = Empty
    Form_Tire_Inventory.txtMountedRim.Value = Empty
    Form_Tire_Inventory.ddnLocation.Value = Empty
    Form_Tire_Inventory.ddnActiv.Value = Empty
    Form_Tire_Inventory.txtTotalHeatCycle.Value = Empty
    Form_Tire_Inventory.txtTotalTimeOnTire.Value = Empty
    Form_Tire_Inventory.txtNote.Value = Empty

End Sub

Sub DeleteTireInventoryData()
    
    Dim strQry As String
    
   strQry = "Delete * from tblTireInventory where ID=" & Form_Tire_Inventory.lblID.Caption
    
    CurrentDb.Execute strQry
    
    Form_SubTireInventory.Requery
    
    MsgBox "Record has been Delected successfully!!!", vbInformation
    
End Sub

Public Function fnPartTireInventorySave(strStatus As String)
    
    Dim rsRecordset As Recordset
    Dim strQry As String
             
    strQry = "Select * from tblPartInventory where ID=" & Form_Part_Tire_Inventory.lblID.Caption
    
    Set rsRecordset = CurrentDb.OpenRecordset(strQry)

    With rsRecordset
    
        If strStatus = "Save" Then
             .AddNew
        Else
              .Edit
        End If
        ![DBR Part Number] = Form_Part_Tire_Inventory.txtDBRPartNo.Value
        ![MFG Part Number] = Form_Part_Tire_Inventory.txtMFGPartNo.Value
        ![Car Type] = Form_Part_Tire_Inventory.txtCarType.Value
        ![Part Type] = Form_Part_Tire_Inventory.txtPartType.Value
        ![Brand] = Form_Part_Tire_Inventory.txtBrand.Value
        ![Model] = Form_Part_Tire_Inventory.txtModel.Value
        ![Description] = Form_Part_Tire_Inventory.txtDescription.Value
        ![Date Acquired] = Form_Part_Tire_Inventory.txtDataAcquired.Value
        ![Date Serviced] = Form_Part_Tire_Inventory.txtDataServiced.Value
        ![Vendor] = Form_Part_Tire_Inventory.txtVendor.Value
        ![Purchase Invoice] = Form_Part_Tire_Inventory.txtPurchaseInvoice.Value
        ![Cost] = Form_Part_Tire_Inventory.txtCost.Value
        ![Condition Purchased] = Form_Part_Tire_Inventory.txtCondPurchased.Value
        ![Current Condition] = Form_Part_Tire_Inventory.txtCurrentCond.Value
        ![Location] = Form_Part_Tire_Inventory.ddnLocation.Value
        ![Active] = Form_Part_Tire_Inventory.ddnActiv.Value
        ![Note] = Form_Part_Tire_Inventory.txtNote.Value
        ![ImgPath] = Form_Part_Tire_Inventory.txtPath.Value
        .Update

    End With
    
     If Form_Part_Tire_Inventory.cmdSave.Caption = "Save" Then
        MsgBox "Record has been saved successfully!!!", vbInformation
    Else
         MsgBox "Record has been updated successfully!!!", vbInformation
    End If
    
    'Call fnInsertSummary
    'Refresh
    Form_SubPartInventory.Requery
    
    Call ResetPartInventory
End Function

Public Function fnInsertSummary()
    
    On Error GoTo ErrInsert
    Dim strQry As String
    strQry = "Insert Into tblSummary([TireID],[Track],RaceDate,Event,BaseCarSetup,Time,Type) Select [TireID],[Track],RaceDate,Event,BaseCarSetup,Time,Type from qrySummary"
    CurrentDb.Execute strQry
    On Error GoTo 0
    
    Exit Function
    
ErrInsert:
    MsgBox Err.Description, vbCritical
    
End Function

Sub ResetPartInventory()
    
    Form_Part_Tire_Inventory.lblID.Caption = 0
    Form_Part_Tire_Inventory.txtDBRPartNo.Value = Empty
    Form_Part_Tire_Inventory.txtMFGPartNo.Value = Empty
    Form_Part_Tire_Inventory.txtCarType.Value = Empty
    Form_Part_Tire_Inventory.txtPartType.Value = Empty
    Form_Part_Tire_Inventory.txtBrand.Value = Empty
    Form_Part_Tire_Inventory.txtModel.Value = Empty
    Form_Part_Tire_Inventory.txtDescription.Value = Empty
    Form_Part_Tire_Inventory.txtDataAcquired.Value = Empty
    Form_Part_Tire_Inventory.txtDataServiced.Value = Empty
    Form_Part_Tire_Inventory.txtVendor.Value = Empty
    Form_Part_Tire_Inventory.txtPurchaseInvoice.Value = Empty
    Form_Part_Tire_Inventory.txtCost.Value = Empty
    Form_Part_Tire_Inventory.txtCondPurchased.Value = Empty
    Form_Part_Tire_Inventory.txtCurrentCond.Value = Empty
    Form_Part_Tire_Inventory.ddnLocation.Value = Empty
    Form_Part_Tire_Inventory.ddnActiv.Value = Empty
    Form_Part_Tire_Inventory.txtNote.Value = Empty
    Form_Part_Tire_Inventory.txtPath.Value = Empty

End Sub

Sub DeletePartInventoryData()
    
    Dim strQry As String
    
   strQry = "Delete * from tblPartInventory where ID=" & Form_Part_Tire_Inventory.lblID.Caption
    
    CurrentDb.Execute strQry
    
    Form_SubPartInventory.Requery
    
    MsgBox "Record has been Delected successfully!!!", vbInformation
    
End Sub

Public Function fnEquipmentMaintenanceSave(strStatus As String)
    
    Dim rsRecordset As Recordset
    Dim strQry As String
             
    strQry = "Select * from tblEquipmentMaintenance where ID=" & Form_Equipment_Maintenence.lblID.Caption
    
    Set rsRecordset = CurrentDb.OpenRecordset(strQry)

    With rsRecordset
    
        If strStatus = "Save" Then
             .AddNew
        Else
              .Edit
        End If
        ![Service_Record] = Form_Equipment_Maintenence.txtServiceRecord.Value
        ![Service_Date] = Form_Equipment_Maintenence.txtDate.Value
        ![Equipment] = Form_Equipment_Maintenence.txtEquipment.Value
        ![Service_Type] = Form_Equipment_Maintenence.txtServiceType.Value
        ![Service_By] = Form_Equipment_Maintenence.txtServiceBy.Value
        ![Description] = Form_Equipment_Maintenence.txtDescription.Value
        ![Notes] = Form_Equipment_Maintenence.lstNote.Value
        .Update

    End With
    
     If Form_Equipment_Maintenence.cmdSave.Caption = "Save" Then
        MsgBox "Record has been saved successfully!!!", vbInformation
    Else
         MsgBox "Record has been updated successfully!!!", vbInformation
    End If

    'Refresh
    Form_SubEquipmentMaintenance.Requery
    
    Call ResetEquipmentMaintenance
End Function

Sub ResetEquipmentMaintenance()
    
   Form_Equipment_Maintenence.lblID.Caption = 0
   Form_Equipment_Maintenence.txtServiceRecord.Value = Empty
   Form_Equipment_Maintenence.txtDate.Value = Empty
   Form_Equipment_Maintenence.txtEquipment.Value = Empty
   Form_Equipment_Maintenence.txtServiceType.Value = Empty
   Form_Equipment_Maintenence.txtServiceBy.Value = Empty
   Form_Equipment_Maintenence.txtDescription.Value = Empty
   Form_Equipment_Maintenence.lstNote.Value = Empty

End Sub

Sub DeleteEquipmentMaintenance()
    
    Dim strQry As String
    
    strQry = "Delete * from tblEquipmentMaintenance where ID=" & Form_Equipment_Maintenence.lblID.Caption
    
    CurrentDb.Execute strQry
    
    Form_SubEquipmentMaintenance.Requery
    
    MsgBox "Record has been Delected successfully!!!", vbInformation
    
End Sub

Public Function fnCheckUpdate(strFieldName As String, lID As Long, strUpdateValue) As Boolean
    
    Dim strPreValue As String
    Dim bUpdate As Boolean
    
    bUpdate = False
    On Error Resume Next
       strUpdateValue = DLookup("& strfieldname &", "tblRaceSession", "ID=" & lID)
          
       If strUpdateValue <> strPreValue Then
            bUpdate = True
            
        ElseIf strUpdateValue = strPreValue Then
            bUpdate = False
        End If
    
        fnCheckUpdate = bUpdate
    On Error GoTo 0

End Function

Public Function fnResetColor()
    
    Dim ctrl As Control
    
    For Each ctrl In Form_Race_Session.Controls
        If InStr(1, ctrl.Name, "txt") > 0 Or InStr(1, ctrl.Name, "ddn") > 0 Then
            If ctrl.BackColor = vbGreen Then
                ctrl.BackColor = vbWhite
            End If
        End If
    Next
    
End Function

Function fnSummaryQuery(strTireID)
    
    Dim qry As QueryDef
    Dim strQry As String
    
    strQry = "Select [Tire ID],[Duration],Type,Track,RaceDate,Event,[Base Car Set],[Time] from tblTireIDCombined where [Tire ID] ='" & strTireID & "'"
    
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "qrySummary"
    Set qry = CurrentDb.CreateQueryDef("qrySummary")
    qry.SQL = strQry
    
    Form_subSummary.Requery
    Form_Tire_Inventory.childSummary.SourceObject = "subSummary"
    
    Form_Tire_Inventory.txtTotalHeatCycle.Value = DCount("[Tire ID]", "tblTireIDCombined", "[Tire ID]='" & strTireID & "'")
    Form_Tire_Inventory.txtTotalTimeOnTire.Value = DSum("Duration", "tblTireIDCombined", "[Tire ID]='" & strTireID & "'")
   
    On Error GoTo 0
    
End Function

