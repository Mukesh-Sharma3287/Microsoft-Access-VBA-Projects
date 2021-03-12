Attribute VB_Name = "ModFunction"
Option Compare Database
Option Explicit

''***********************************************************************************************************
'' Form Name        : ModFunction
'' Description      : This Module contain the Various function to use the entire tool
''
''
''    Date                  Developer               Created             Remarks
''  10-Jan-2021            Mukesh Sharma             First Version
''***********************************************************************************************************

''******************************************************************''
''  Function Name      :  UpdateLinkTables                        ''
''  Description        :  Update the links of the linked tables  ''
''  Arguments          :  Na                                      ''
''  Return              :  Na                                      ''
''******************************************************************''
Sub UpdateLinkTables(strPath As String)
    
    Dim dbDatabase As DAO.Database
    Dim tdf As TableDef
    
    Set dbDatabase = CurrentDb
    
    'Update the new path of the link tables
    For Each tdf In dbDatabase.TableDefs
  
       If tdf.SourceTableName <> "" And tdf.Name <> "tblTaxiDetails" Then
            tdf.Properties("Attributes").Value = 0
            tdf.Connect = ";DATABASE=" & strPath & "; PWD=" & gstrDBPassword
            tdf.RefreshLink
            Application.SetHiddenAttribute acTable, tdf.Name, True
            
            'Make the table as HiddenObject
            tdf.Properties("Attributes").Value = dbHiddenObject
            'Make the table as normal object
            'tdf.Properties("Attributes").Value = 0
        End If
    Next
    
    'Close the objects
    Set tdf = Nothing
    Set dbDatabase = Nothing

End Sub


''******************************************************************''
''  Function Name      :  fnToCheckFolderExists                  ''
''  Description        :  To check given file/folder is available ''
''                          in the system or not.                  ''
''  Arguments          :  strFullPath as String - To get file path''
''  Return              :  Boolean - True means is available.    ''
''                          FALSE means not available.            ''
''******************************************************************''
Function fnToCheckFolderExists(strFullPath As String) As Boolean
    
    fnToCheckFolderExists = False
    On Error GoTo ErrExit
    
    'Checks the File/Folder available or not.
    If Dir(strFullPath, vbDirectory) <> "" Then fnToCheckFolderExists = True
    
ErrExit:

    On Error GoTo 0
    
End Function

''***********************************************************************************************************
'' Procedure Name       : fnUserName
'' Description          : To Extract Current User Name
'' Arguments            : NA
'' Returns              : NA
''***********************************************************************************************************

Public Function fnUserName()

fnUserName = VBA.Environ("UserName")

End Function


''***********************************************************************************************************
'' Procedure Name       : fnCheck_File_Exists
'' Description          : To Check existing file
'' Arguments            : ByVal strPath As String
'' Returns              : Boolean: Returns True If file is exist else False
''***********************************************************************************************************

Public Function fnCheck_File_Exists(ByVal strPath As String) As Boolean

    If Dir(Trim(strPath)) <> "" Then
        fnCheck_File_Exists = True
    Else
        fnCheck_File_Exists = False
    End If
    
End Function

