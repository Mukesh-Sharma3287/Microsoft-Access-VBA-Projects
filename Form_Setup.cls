VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdDelete_Click()
    Call ModSave.DeleteSetupData
End Sub

Private Sub cmdReset_Click()
    Call ModSave.ResetSetup
End Sub

Private Sub cmdSave_Click()
    Call ModSave.submitCarSetup("Save")
End Sub

Private Sub cmdUpdate_Click()
    Call ModSave.submitCarSetup("Update")
End Sub
