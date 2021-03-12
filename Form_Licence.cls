VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Licence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

''***********************************************************************************************************
'' Form Name        : Form_Licence
'' Description      : This Form contain the procedure to Licence
''                    with Databse path
''
''    Date                  Developer               Created             Remarks
''  11-Nov-2016            Mukesh Sharma             First Version
''***********************************************************************************************************

''***********************************************************************************************************
'' Procedure Name       : Form_Load
'' Description          : To Initialize the userid
'' Arguments            : NA
'' Returns              : NA
''***********************************************************************************************************

Private Sub Form_Load()

    Form_Licence.txtUsrID.Value = Environ("UserName")
    Form_Licence.txtUsrID.Enabled = False

End Sub

''***********************************************************************************************************
'' Procedure Name       : cmdCance_Click
'' Description          : To Cancel the form
'' Arguments            : NA
'' Returns              : NA
''***********************************************************************************************************

Private Sub cmdCance_Click()

    DoCmd.Close acForm, "Licence", acSaveYes

End Sub

''***********************************************************************************************************
'' Procedure Name       : cmdSubmit_Click
'' Description          : To Add the Licence details
'' Arguments            : NA
'' Returns              : NA
''***********************************************************************************************************

Private Sub cmdSubmit_Click()

    Dim rsRecordset As Recordset
    
    'Validation for details
    If Form_Licence.txtExpireson.Value = "" Or IsNull(Form_Licence.txtExpireson.Value) = True Or IsDate(Form_Licence.txtExpireson.Value) = False _
        Or Form_Licence.txtLicencePwd.Value = "" Or IsNull(Form_Licence.txtLicencePwd.Value) = True Then
        MsgBox "Please fill all details to proceed", vbCritical, Tool_Name
        Exit Sub
    End If
    
    'Validation for password
    If Form_Licence.txtLicencePwd.Value <> Chr(115) & Chr(116) & Chr(97) & Chr(114) & Chr(108) & Chr(107) & Chr(103) Then
        MsgBox "Password is Not Valid ,Please try agian", vbCritical, Tool_Name
        Form_Licence.txtExpireson.Value = ""
        Form_Licence.txtLicencePwd.Value = ""
        Exit Sub
    End If
    
    'Delete the data from table
    CurrentDb.Execute ("Delete * from tblLicence")
    
    Set rsRecordset = CurrentDb.OpenRecordset("Select * from tblLicence")
        
    With rsRecordset
        .AddNew
        !UserID = Form_Licence.txtUsrID.Value
        !Expireson = Form_Licence.txtExpireson.Value
        !LicenceKeys = Form_Licence.txtLicencePwd.Value
        .Update
    End With
    
    Form_Licence.txtExpireson.Value = ""
    Form_Licence.txtLicencePwd.Value = ""
    
    MsgBox "Licence Details Added Successfully!!", vbInformation, Tool_Name
    
    DoCmd.OpenForm "Login", acNormal
    DoCmd.Close acForm, "Licence", acSaveYes
    
    Set rsRecordset = Nothing

End Sub


