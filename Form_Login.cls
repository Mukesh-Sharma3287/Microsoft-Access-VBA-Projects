VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Compare Database
'Option Explicit
'
'''***********************************************************************************************************
''' Form Name        : Form_Login
''' Description      : This Form contain the procedure to Login
'''                    with Databse path
'''
'''    Date                  Developer               Created             Remarks
'''  11-Nov-2016            Mukesh Sharma             First Version
'''***********************************************************************************************************
'
'''***********************************************************************************************************
''' Procedure Name       : Form_Load
''' Description          : To Initialize the data
''' Arguments            : NA
''' Returns              : NA
'''***********************************************************************************************************
'
Private Sub Form_Load()

    Dim rsRecordset As Recordset

    Application.SetOption "Show Hidden Objects", False

    DoCmd.LockNavigationPane True


    'Disable right clicks and menus
    CurrentDb.Properties("AllowFullMenus") = False
    CurrentDb.Properties("AllowShortcutMenus") = False
    CurrentDb.Properties("AllowBuiltinToolbars") = False


    'Check that database file is present at the set location

    Set rsRecordset = CurrentDb.OpenRecordset("Select Database_Location from tblDatabaseLocation")

    With rsRecordset
        If .EOF = True Then

            'If rsRecordset is coming as blank then first insert a new row
            CurrentDb.Execute ("Insert into tblDatabaseLocation (Database_Location) Values ('Select the database')")

             'Show the prompt and ask the user for select the database file
            MsgBox "It seems that Database path is not present in database,Please add the database path", vbCritical, Tool_Name
            DoCmd.OpenForm "Database", acNormal
            DoCmd.Close acForm, "Login", acSaveYes
            Exit Sub
        ElseIf InStr(rsRecordset!Database_Location, "Micra Cup") = 0 Then

            MsgBox "It Seems wrong database is selected." & vbNewLine & "Please select proper database to run this tool", vbInformation, Tool_Name
            DoCmd.OpenForm "Database", acNormal
            DoCmd.Close acForm, "Login", acSaveYes
            Exit Sub
        End If

        If ModFunction.fnToCheckFolderExists(rsRecordset!Database_Location) = False Then
            'Show the prompt and ask the user for select the database file
            MsgBox "It seems the database file is not present in the set location" & vbNewLine & _
            "Please update the database location to proceed", vbInformation, Tool_Name

            DoCmd.OpenForm "Database", acNormal
            DoCmd.Close acForm, "Login", acSaveYes
            Exit Sub
        End If


    End With

    Form_Login.ddnUserID.Value = VBA.Environ("UserName")
    Form_Login.txtPwd.Value = ""
    Form_Login.ddnUserID.SetFocus
'
    Form_Login.ddnUserID.Value = fnUserName
'
   Set rsRecordset = Nothing


End Sub

''***********************************************************************************************************
'' Procedure Name       : cmdLogin_Click
'' Description          : To Login the form
'' Arguments            : NA
'' Returns              : NA
''***********************************************************************************************************

Private Sub cmdLogin_Click()

    Dim rsRecordset As Recordset

    'Validation
    If Form_Login.ddnUserID.Value = "" Or IsNull(Form_Login.ddnUserID.Value) = True _
        Or Form_Login.txtPwd.Value = "" Or IsNull(Form_Login.txtPwd.Value) = True Then
        MsgBox "Please fill all details to Proceed", vbCritical, Tool_Name
        Exit Sub
    End If

    Set rsRecordset = CurrentDb.OpenRecordset("Select * from tblUser where UserID='" & Form_Login.ddnUserID.Value & "'")

    With rsRecordset

        'Validation for User ID
        If .EOF = True Then
            MsgBox "your user ID is not present in Database, make sure you have to right this tool", vbCritical, Tool_Name
            Form_Login.txtPwd.Value = ""
            Exit Sub
        End If

        'Validation for Password
        If !Password <> Form_Login.txtPwd.Value Then
            MsgBox "Password is Incorrect, Please try Again", vbCritical, Tool_Name
            Form_Login.txtPwd.Value = ""
            Exit Sub
        End If


    End With

    MsgBox "Login Successfully!!", vbInformation, Tool_Name

    gstrUserID = Form_Login.ddnUserID.Value

    DoCmd.OpenForm "Home", acNormal
    DoCmd.Maximize
    DoCmd.Close acForm, "Login", acSaveYes

     Set rsRecordset = Nothing

End Sub

''***********************************************************************************************************
'' Procedure Name       : cmdCancel_Click
'' Description          : To Cancel the form
'' Arguments            : NA
'' Returns              : NA
''***********************************************************************************************************

Private Sub cmdCancel_Click()

    DoCmd.Close acForm, "Login", acSaveYes

End Sub
