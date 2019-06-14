VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCTmplateCoversheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const mstrPATH_TEMPLATE As String = "C:\Documents and Settings\ajones\Application Data\Microsoft\Templates\Contracts Management\"
Private Const mstrTEMPLATE_NAME As String = "Coversheet for Signature v6.dotx"


Private Const mstrDATA_FILE_XLS As String = "Coversheet Data.xlsx"
Private Const mstrSHAREPOINT_LOCATION = "https://mysharepoint/sdp/oosvp/RAS/Contracts Management/CM-transfer/Contracts Management/Contracts Team/Activities/Reports/Data Links for Reports"
Private Const mstrWORKSHEET_TAB = "Coversheet Data"
Private Const mstrPATH_DATASOURCE = "CBDataSource"

Private mobjCoversheetDocument As Word.Document

Private Sub Class_Initialize()
'Class initialize
'Class should make connection to the data file and if it doesn't, raise an error
    On Error GoTo ErrorHandler
    
'    ConnectToMergeData
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Source = "Class_Initialize"
        .Description = "Class clsCTmplateCoversheet initialization failed"
        .Raise .Number, .Source, .Description
    End With
End Sub

Private Sub ConnectToMergeData()
    Const strNETWORK_DRIVE_SHAREPOINT_LOCATION As String = "I:\Coversheet - Salesforce Export.xlsx"
    Const strWORKBOOK As String = "Coversheet Data.xlsx"
    Const strWORKSHEET_TAB As String = "Contracts List"
    
    On Error GoTo ErrorHandler

    'May want to use an ODC file and pass that to the code below instead of a hardcoded path or url.
    'NOTE, that file may have to be pushed downed to each users machine though, WHICH is the same as
    'pushing the actual data file down to each user.
    'Microsoft says I can't use Http:// Article => https://support.office.com/en-us/article/Data-sources-you-can-use-for-a-mail-merge-9de322a6-f0f9-448d-a113-5fab317d9ef4

    ActiveDocument.MailMerge.OpenDataSource Name:=Environ(mstrPATH_DATASOURCE) & "\" & strWORKBOOK, _
        ReadOnly:=True, _
        Connection:=strWORKSHEET_TAB

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Connect To Merge Data"
    End With
    Resume Exit_Here
End Sub

Public Sub CreateCoversheetForSignatureForSpecificContract(ContractNumber As String)
    Dim dsMain As MailMergeDataSource
    Dim numRecord As Integer
    Const strMergeField As String = "Contract_Number"
    
    On Error GoTo ErrorHandler
    
    ActiveDocument.MailMerge.ViewMailMergeFieldCodes = False
    Set dsMain = ActiveDocument.MailMerge.DataSource
    dsMain.ActiveRecord = wdFirstRecord
    If dsMain.FindRecord(FindText:=ContractNumber, Field:=strMergeField) = True Then
        numRecord = dsMain.ActiveRecord
    End If

Exit_Here:
    Set dsMain = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Create Coversheet For Signature For Specific Contract"
    End With
    Resume Exit_Here
End Sub

Public Property Get CoversheetTemplateName()
    On Error GoTo ErrorHandler
    
    CoversheetTemplateName = mstrPATH_TEMPLATE & mstrTEMPLATE_NAME
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Source = "Get CoversheetTemplateName"
        .Description = "Recovering Template Name failed"
        .Raise .Number, .Source, .Description
    End With
End Property

Private Sub Class_Terminate()
'Class terminate
    On Error GoTo ErrorHandler
    
    
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Source = "Class_Terminate"
        .Description = "Class clsCTmplateCoversheet Termination failed"
        .Raise .Number, .Source, .Description
    End With
End Sub

Public Property Get SharePointLocation()
    On Error GoTo ErrorHandler
    
    
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Source = "Get SharePointLocation"
        .Description = "Recovering SharePointLocation failed"
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Get ExcelDataFile()
    On Error GoTo ErrorHandler
    
    
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Source = "Get ExcelDataFile"
        .Description = "Recovering Name of the Excel Data File failed"
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Get ExcelWorksheetTabName()
    On Error GoTo ErrorHandler
    
    
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Source = "Get ExcelWorksheetTabName"
        .Description = "Recovering Name of the Excel Worksheet Tab failed"
        .Raise .Number, .Source, .Description
    End With
End Property




