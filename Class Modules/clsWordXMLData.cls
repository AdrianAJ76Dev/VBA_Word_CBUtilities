VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWordXMLData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrDefaultPath As String
Private mstrXMLPath As String
Private mstrXMLFile As String
Private mstrXMLFullName As String
Private mintCustomDataPart As Integer

'December 5, 2013
'Can load and reference more than one XML file
Private mvarXMLDataFiles As Variant
Private mvarXMLDataFile As Variant

Private Const intCUSTOM_DATA As Integer = 4
Private Const strPATH_PIVOTAL_EXPORT_XML As String = "C:\Documents and Settings\ajones\Desktop\"

'Constants - XML Fields - Quote
'Coming Soon!

Private Const mstrPARENT_NAME As String = "/Results_List/Rows/Row/"

'Constants - XML Fields - Contact
Private Const mstrContactFIRST_NAME As String = "First_Name"
Private Const mstrContactLAST_NAME As String = "Last_Name"
Private Const mstrContactFULL_NAME As String = "Full_Name"
Private Const mstrContactJOB_TITLE As String = "Job_Title"
Private Const mstrCLIENT As String = "Company_Id_Rn_Descriptor"
Private Const mstrEM As String = "Account_Manager_Id_Rn_Descriptor"
Private Const mstrPHONE As String = "CB_Direct_Phone"
Private Const mstrCITY As String = "City"
Private Const mstrADDRESS_1 As String = "Address_1"
Private Const mstrADDRESS_2 As String = "Address_2"
Private Const mstrEMAIL As String = "Email"
Private Const mstrZIP As String = "Zip"

'Constants - XML Fields - Education Manager
Private Const mstrEM_FIRST_NAME As String = "First_Name"
Private Const mstrEM_LAST_NAME As String = "Last_Name"
Private Const mstrEM_JOB_TITLE As String = "Job_Title"
Private Const mstrEM_PHONE As String = "Work_Phone"
Private Const mstrEM_EMAIL As String = "Work_Email"

'Constants - XML Fields - Opportunity
Private Const mstrOP_REVENUE_NO As String = "CB_Contr_Rev_Num"
Private Const mstrOP_PRODUCTS As String = "CB_Product_Interest"

Private Enum ccFields
    Client
    EMFirstName
    
End Enum

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    mstrDefaultPath = strPATH_PIVOTAL_EXPORT_XML

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Raise .Number & vbCr & .Description, "Property Get XMLContactInfo"
    End With
End Sub

Public Property Get XMLContactInfo() As String
    On Error GoTo ErrorHandler
    
    XMLContactInfo = mstrXMLFile
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number & vbCr & .Description, "Property Get XMLContactInfo"
    End With
End Property

Public Property Get DefaultPath() As String
    On Error GoTo ErrorHandler
    
    DefaultPath = mstrDefaultPath
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number & vbCr & .Description, "Property Get XMLContactInfo"
    End With
End Property

Public Property Let XMLFiles(ByVal vvarFileNames As Variant)
    On Error GoTo ErrorHandler
    
    mvarXMLDataFiles = vvarFileNames
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number & vbCr & .Description, "Property Let XMLFile"
    End With
End Property

Public Property Let XMLFile(ByVal vstrXMLFileFullName As String)
    On Error GoTo ErrorHandler
    
    mstrXMLFullName = vstrXMLFileFullName
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number & vbCr & .Description, "Property Let XMLFile"
    End With
End Property

Public Property Get XMLFile() As String
    On Error GoTo ErrorHandler
    
    XMLFile = mstrXMLFullName
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number & vbCr & .Description, "Property Get XMLFile"
    End With
End Property

Public Sub BindContentControls()
    On Error GoTo ErrorHandler
    
    If Not ActiveDocument.Type = wdTypeTemplate Then
        
        mintCustomDataPart = ActiveDocument.CustomXMLParts.Count + 1
        ActiveDocument.CustomXMLParts.Add
        ActiveDocument.CustomXMLParts(mintCustomDataPart).Load mstrXMLFullName
        
'        With ActiveDocument
'            .ContentControls(Client).XMLMapping.SetMapping mstrPARENT_NAME & mstrCLIENT
'            .ContentControls(FirstName).XMLMapping.SetMapping mstrPARENT_NAME & mstrContactFIRST_NAME
'            .ContentControls(LastName).XMLMapping.SetMapping mstrPARENT_NAME & mstrContactLAST_NAME
'            .ContentControls(JobTitle).XMLMapping.SetMapping mstrPARENT_NAME & mstrContactJOB_TITLE
'            .ContentControls(EM).XMLMapping.SetMapping mstrPARENT_NAME & mstrEM
'        End With
    End If

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Raise .Number & vbCr & .Description, "BindContentControls"
    End With
End Sub

Public Sub BindContentControlsv2()
    Dim ccCtrl As ContentControl
    On Error GoTo ErrorHandler
    
    If Not ActiveDocument.Type = wdTypeTemplate Then
        mintCustomDataPart = ActiveDocument.CustomXMLParts.Count + 1
        ActiveDocument.CustomXMLParts.Add
        ActiveDocument.CustomXMLParts(mintCustomDataPart).Load mstrXMLFullName
        
        With ActiveDocument
            For Each ccCtrl In .SelectContentControlsByTitle("Client")
                ccCtrl.XMLMapping.SetMapping mstrPARENT_NAME & mstrCLIENT
            Next ccCtrl
            
            For Each ccCtrl In .SelectContentControlsByTitle("First Name")
                ccCtrl.XMLMapping.SetMapping mstrPARENT_NAME & mstrContactFIRST_NAME
            Next ccCtrl
            
            For Each ccCtrl In .SelectContentControlsByTitle("Last Name")
                ccCtrl.XMLMapping.SetMapping mstrPARENT_NAME & mstrContactLAST_NAME
            Next ccCtrl
            
            For Each ccCtrl In .SelectContentControlsByTitle("Job Title")
                ccCtrl.XMLMapping.SetMapping mstrPARENT_NAME & mstrContactJOB_TITLE
            Next ccCtrl
            
            For Each ccCtrl In .SelectContentControlsByTitle("EM")
                ccCtrl.XMLMapping.SetMapping mstrPARENT_NAME & mstrEM
            Next ccCtrl
        End With
    Else
        MsgBox ActiveDocument.Name & " is a template.  Please select or create a document based on this template: " & ActiveDocument.AttachedTemplate.Name
    End If

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Raise .Number & vbCr & .Description, "BindContentControls"
    End With
End Sub
