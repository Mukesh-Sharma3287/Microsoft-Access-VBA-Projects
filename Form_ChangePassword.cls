VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

''***********************************************************************************************************
'' Form Name        : Form_ChangePassword
'' Description      : This Form contain the procedure to Change the password
''
''
''    Date                  Developer               Created             Remarks
''  11-Nov-2016            Mukesh Sharma             First Version
''***********************************************************************************************************

''***********************************************************************************************************
'' Procedure Name       : cmdOK_Click
'' Description          : To Change the Old password
'' Arguments            : NA
'' Returns              : NA
''***********************************************************************************************************

Private Sub cmdOK_Click()

    Dim rsRecordset As Recordset
        
    'Validation
    If Form_ChangePassword.txtOldPass.Value = "" Or IsNull(Form_ChangePassword.txtOldPass.Value) = True _
        Or Form_ChangePassword.txtNewPass.Value = "" Or IsNull(Form_ChangePassword.txtNewPass.Value) = True Then
        MsgBox "Please Fill the Details to Proceed", vbCritical, Tool_Name
        Exit Sub
    End If
    
    Set rsRecordset = CurrentDb.OpenRecordset("Select * From tblUser Where UserID='" & gstrUserID & "'")
    
    With rsRecordset
        
        If !Password <> Form_ChangePassword.txtOldPass.Value Then
            MsgBox "Old Password is not match,Please Try again", vbCritical, Tool_Name
            Form_ChangePassword.txtNewPass.Value = ""
            Form_ChangePassword.txtOldPass.Value = ""
            Exit Sub
        End If
            
        CurrentDb.Execute ("Update tblUser Set Password='" & Form_ChangePassword.txtNewPass.Value & "'" & " Where UserID='" & gstrUserID & "'")
        MsgBox "Old Password has been Changed!!", vbInformation, Tool_Name
        Form_ChangePassword.txtNewPass.Value = ""
        Form_ChangePassword.txtOldPass.Value = ""
        DoCmd.Close acForm, "ChangePassword", acSaveYes
    
    End With
    
    Set rsRecordset = Nothing

End Sub

''***********************************************************************************************************
'' Procedure Name       : cmdCancel_Click
'' Description          : To Cancel the form
'' Arguments            : NA
'' Returns              : NA
''***********************************************************************************************************

Private Sub cmdCancel_Click()

    DoCmd.Close acForm, "ChangePassword", acSaveYes

End Sub
