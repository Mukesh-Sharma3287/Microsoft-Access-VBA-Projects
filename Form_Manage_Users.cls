VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Manage_Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdAddNew_Click()
    
    'Call the function to
    Call ModSett.fnAddNew
    
End Sub

Private Sub cmdReset_Click()
    
  
       
End Sub

Private Sub cmdClear_Click()
    Form_Manage_Users.txtUsrID.Value = ""
    Form_Manage_Users.txtUsrName.Value = ""
    Form_Manage_Users.txtPass.Value = ""
    Form_Manage_Users.ddnRole.Value = ""
    Form_Manage_Users.cmdAddNew.Caption = "Add New"
    Form_Manage_Users.txtUsrID.Locked = False
End Sub

Private Sub cmdDelete_Click()
    
    Dim strResponse As String
    'Validation
    If Form_Manage_Users.txtUsrID.Value = "" Or IsNull(Form_Manage_Users.txtUsrID.Value) = True Then
        MsgBox "Please Select the UserID to delete", vbInformation
    End If
    
    strResponse = MsgBox("Are you sure to Delete record?", vbYesNo)
    
    If strResponse = vbYes Then
        CurrentDb.Execute ("Delete * from tblUser Where UserID='" & Form_Manage_Users.txtUsrID.Value & "'")
        Form_SubUsrDetails.Requery
    ElseIf strResponse = vbNo Then
        Exit Sub
    End If
    
    
End Sub

Private Sub Form_Current()

End Sub
