VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ImportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

''***********************************************************************************************************
'' Form Name        : Form_ImportData
'' Description      : This Form contain the procedure to browse the File path and import the data
''
''
''    Date                  Developer               Created             Remarks
''  11-Nov-2016            Mukesh Sharma             First Version
''***********************************************************************************************************

''***********************************************************************************************************
'' Procedure Name       : cmdBrowse_Click
'' Description          : To Browse the Import data file path
'' Arguments            : NA
'' Returns              : NA
''***********************************************************************************************************

Private Sub cmdBrowse_Click()

 'Declare the Variable
    Dim objFileDialog As FileDialog
    Dim vSelected As Variant
    
    'Set the Object
    Set objFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    'Filter the files
    With objFileDialog
        .AllowMultiSelect = False
        .ButtonName = "Select File"
        .Filters.Clear
        .Filters.Add "Access File", "*.Accdb;*.mddb", 1
        .Title = "Select Database path"
        .Show
        
        For Each vSelected In objFileDialog.SelectedItems
            Form_Database.txtPath.Value = vSelected
        Next
    
    End With
    
    'Clear the memory
    Set objFileDialog = Nothing


End Sub


''***********************************************************************************************************
'' Procedure Name       : cmdCancel_Click
'' Description          : To Cancel the form
'' Arguments            : NA
'' Returns              : NA
''***********************************************************************************************************

Private Sub cmdCancel_Click()

    DoCmd.Close acForm, "ImportData", acSaveYes

End Sub

''***********************************************************************************************************
'' Procedure Name       : cmdOK_Click
'' Description          : To Link the all table with database
'' Arguments            : NA
'' Returns              : NA
''***********************************************************************************************************

Private Sub cmdOK_Click()

    Dim rsRecordset As Recordset

    'Validation
    If Form_Database.txtPath.Value = "" Or IsNull(Form_Database.txtPath.Value) = True Then
        MsgBox "Please Select the Database to Proceed", vbCritical, Tool_Name
        Exit Sub
    End If
    
    'Validation
    If ModFunction.fnCheck_File_Exists(Form_Database.txtPath.Value) = False Then
        MsgBox "File is Not exist,Please Select the Valid path of Database", vbCritical, Tool_Name
        Exit Sub
    End If
    
    Set rsRecordset = CurrentDb.OpenRecordset("Select Top 1 * from tblDatabaseLocation")
    
    With rsRecordset
        If .EOF = True Then
            CurrentDb.Execute ("Insert Into tblDatabaseLocation (Database_Location) Values('Select Database Location')")
            MsgBox "It Seems that Database Location is not present Database,Please Select the Database Location", vbCritical, Tool_Name
            Exit Sub
        End If
        
        CurrentDb.Execute ("Update tblDatabaseLocation Set Database_Location='" & Form_Database.txtPath.Value & "'")
    
    End With
            
    'Call the function to Link the table
    ModFunction.fnLinkTable (Form_Database.txtPath.Value)
    

End Sub

