Attribute VB_Name = "ModSett"
Option Compare Database
Option Explicit

Public Function fnAddNew()
    
    Dim rsRecordset As Recordset
    Dim strQry As String
    
    If Form_Manage_Users.cmdAddNew.Caption = "Add New" Then
         
         'Validation
         If Form_Manage_Users.txtUsrID.Value = "" Or IsNull(Form_Manage_Users.txtUsrID.Value) = True Then
            MsgBox "Please Enter the UserID to proceed", vbCritical
        ElseIf Form_Manage_Users.txtUsrName.Value = "" Or IsNull(Form_Manage_Users.txtUsrName.Value) = True Then
            MsgBox "Please Enter the UserName to proceed", vbCritical
        ElseIf Form_Manage_Users.txtPass.Value = "" Or IsNull(Form_Manage_Users.txtPass.Value) = True Then
            MsgBox "Please Enter the Password to proceed", vbCritical
        ElseIf Form_Manage_Users.ddnRole.Value = "" Or IsNull(Form_Manage_Users.ddnRole.Value) = True Then
            MsgBox "Please Select Role to proceed", vbCritical
        End If
            
         'Validation
         If DCount("UserID", "tblUser", "UserID='" & Form_Home.lblUserName.Caption & "'") = 2 Then
            MsgBox "User Name Should be Unique", vbCritical
        End If
   End If
   
    strQry = "Select * from tblUser where UserID='" & Form_Manage_Users.txtUsrID.Value & "'"
    
    Set rsRecordset = CurrentDb.OpenRecordset(strQry)
    
    With rsRecordset
        If .EOF = False Then
            .Edit
            
        Else
            .AddNew
        End If
        
        !UserID = Form_Manage_Users.txtUsrID.Value
        !UserName = Form_Manage_Users.txtUsrID.Value
        !Password = Form_Manage_Users.txtPass.Value
        !Role = Form_Manage_Users.ddnRole.Value
        .Update
    End With
    
     If Form_Manage_Users.cmdAddNew.Caption = "AddNew" Then
        MsgBox "Record New Added successfully!!!", vbInformation
    Else
        MsgBox "Record Updated successfully!!!", vbInformation
        
    End If
    
     Form_SubUsrDetails.Requery
    
End Function

