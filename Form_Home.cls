VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

''***********************************************************************************************************
'' Form Name        : Form_Home
'' Description      : This Form contain the procedure of Home form
''
''
''    Date                  Developer               Created             Remarks
''  11-Nov-2016            Mukesh Sharma             First Version
''***********************************************************************************************************

''***********************************************************************************************************
'' Procedure Name       : Form_Load
'' Description          : To Initialize the User Name
'' Arguments            : NA
'' Returns              : NA
''***********************************************************************************************************
'
Private Sub Form_Load()

    Dim rsRecordset As Recordset

    Set rsRecordset = CurrentDb.OpenRecordset("Select * from tblUser where UserID='" & gstrUserID & "'")
    
    On Error Resume Next
    With rsRecordset
        lblUserName.Caption = !UserName
    End With
    On Error GoTo 0
    Form_Home.chldHome.SourceObject = "Race_Session"
    
    

    Set rsRecordset = Nothing

End Sub

Private Sub cmdExit_Click()
    
    Application.Quit
    
End Sub

Private Sub cmdCarMaintenence_Click()
    Form_Home.chldHome.SourceObject = "Car_Maintenence"
    Form_Home.lblName.Caption = "Car_Maintenence"
End Sub

Private Sub cmdCarSetup_Click()
    Form_Home.chldHome.SourceObject = "Setup"
    Form_Home.lblName.Caption = "Setup"
End Sub

Private Sub cmdEquipmentMaintenence_Click()
    Form_Home.chldHome.SourceObject = "Equipment_Maintenence"
    Form_Home.lblName.Caption = "Equipment_Maintenence"
End Sub

Private Sub cmdPartInventory_Click()
    Form_Home.chldHome.SourceObject = "Part_Tire_Inventory"
    Form_Home.lblName.Caption = "Part Inventory"
End Sub

Private Sub cmdRaceSession_Click()
    Form_Home.chldHome.SourceObject = "Race_Session"
    Form_Home.lblName.Caption = "Race_Session"
End Sub

Private Sub cmdTireInventory_Click()
    Form_Home.chldHome.SourceObject = "Tire_Inventory"
    Form_Home.lblName.Caption = "Tire_Inventory"
End Sub


Private Sub cmdManageUser_Click()

    'Validation
  If DLookup("Role", "tblUser", "UserID='" & Form_Home.lblUserName.Caption & "'") = "User" Then
    MsgBox "Role '" & Form_Home.lblUserRole.Caption & "' does not have enough rights for this activity", vbInformation, Tool_Name
    Form_Home.chldHome.SourceObject = ""
    Exit Sub
 End If

    Form_Home.chldHome.SourceObject = "Manage_Users"
     Form_Home.lblName.Caption = "Manage_Users"

End Sub

Private Sub cmdReport_Click()
    
    Form_Home.chldHome.SourceObject = "Reports_Menu"
    Form_Home.lblName.Caption = "Reports_Menu"
    
End Sub



''***********************************************************************************************************
'' Procedure Name       : Form_Resize
'' Description          : To Resize the Form
'' Arguments            : NA
'' Returns              : NA
''***********************************************************************************************************

Private Sub Form_Resize()

    If Form_Home.InsideWidth > 16000 Then
        imgHome.Width = Form_Home.InsideWidth - 600
        cmdHome.Left = Form_Home.InsideWidth - 600
    End If

    If Form_Home.InsideWidth < 16000 Then
        imgHome.Width = Form_Home.InsideWidth - 600
        cmdHome.Left = Form_Home.InsideWidth - 600
    End If
    
        
End Sub

''***********************************************************************************************************
'' Procedure Name       : cmdChangePassword_Click
'' Description          : To open form of change password
'' Arguments            : NA
'' Returns              : NA
''***********************************************************************************************************

Private Sub cmdChangePassword_Click()

DoCmd.OpenForm "ChangePassword", acNormal
Form_Home.lblName.Caption = "ChangePassword"

End Sub


