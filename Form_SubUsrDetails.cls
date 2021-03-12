VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SubUsrDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()
    
    On Error Resume Next
    Form_Manage_Users.cmdAddNew.Caption = "Save"
    Form_Manage_Users.txtUsrID.Locked = True
    Form_Manage_Users.txtUsrID.Value = Form_SubUsrDetails.UserID.Value
    Form_Manage_Users.txtUsrName.Value = Form_SubUsrDetails.UserName.Value
    Form_Manage_Users.txtPass.Value = Form_SubUsrDetails.Password.Value
    Form_Manage_Users.ddnRole.Value = Form_SubUsrDetails.Role.Value
    On Error GoTo 0
    
End Sub
