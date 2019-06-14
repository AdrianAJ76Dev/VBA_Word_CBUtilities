VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCoverLetterVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NOTE: Change in class from ThisDocument to Activedocument or a object variable
'Make changes from implicintly referencing the template to explicitly referencing the template.

Option Explicit

Private mstrContact As String
Private mstrTitle As String
Private mstrSchoolDistrict As String
Private mstrProgram As String
Private mstrAddress As String

'4/24/2013
Private mintCopies As Integer

'11/13/2012
Private mstrSignatory As String

'Contract Specialist
Public Enum CSAutoText
    Nicole = 0
    Adrian = 1
    Alex = 2
End Enum

Public Enum ContactAutoText
    ContactTitle = 0
    ContactNoTitle = 1
End Enum

Public Enum LetterType
    FullyExecuted = 0
    Signature = 1
End Enum

'11/13/2012
Private mintLetterType As LetterType

Public Enum ProgramAutoText
    PN = 0      'PSAT
    PN8_9 = 1   'PN 8/9
    PNRS = 2    'PSAT and Readistep
    Path = 3    'Pathway (SAT, PSAT, and ReadiStep)
    PD = 4      'Professional Development
    AP = 5      'Advanced Placement
    APPD = 6    'Advanced Placement Professional Development
    SAT = 7     'SAT
    SB = 8      'Springboard
    MS = 9      'Multiservice (AP Exams, PSAT, and SAT)
    Other = 10  'User will enter program manually
End Enum

Private mstrProgramAutoText(ProgramAutoText.PN To ProgramAutoText.Other) As String

Private mstrContractSpecialistPhoneNum(CSAutoText.Nicole To CSAutoText.Alex) As String
Private mstrContractSpecialistAutoText(CSAutoText.Nicole To CSAutoText.Alex) As String
Private mstrContactInfoAutoText(ContactAutoText.ContactTitle To ContactAutoText.ContactNoTitle) As String

Private mintActiveContractSpecialist As CSAutoText
Private mintContactInfoChoice As ContactAutoText
Private mintProgramChoice As ProgramAutoText

'09/24/2012
'Path currently is hard coded.  This is bad design.  I need to dynamically retreive the path of Building Blocks.dotx
'for currently logged in user (since this document clearly seems to exist under the currently logged in user's directory)
'Windows 7
'Private Const mstrPATH_CBAUTOTEXT As String = "C:\Users\ajones\AppData\Roaming\Microsoft\Document Building Blocks\1033\14\Building Blocks.dotx"
'Windows XP
Private Const mstrPATH_CBAUTOTEXT As String = "C:\Documents and Settings\ajones\Application Data\Microsoft\Document Building Blocks\1033\14\Building Blocks.dotx"

'11/09/2012
Private Const mstrPATH_SOLE_SOURCE_LETTERS As String = "Q:\RAS Contracts Management\Gov Forms and Docs\Sole Source Letters\K12\"

'11/13/2012
'AutoText Constants
Private Const mstrAUTOTEXT_SIGNATURE As String = "Signature - General"
Private Const mstrAUTOTEXT_FULLY_EXECUTED As String = "Fully Executed"

'11/27/2012
Private mstrFilePath As String
Private mstrFileName As String


'07/28/2016
Private Const mstrPATH_SOLE_SOURCE_LETTERS_v2 As String = "C:\Users\ajones\Documents\Uploads to Salesforce\Sole Source Letter Requests\"


Private Const mstrBMK_CLOSING As String = "Closing"
Private Const mstrBMK_CONTACT_INFO As String = "ContactInfo"
Private Const mstrBMK_PROGRAM As String = "Program"
Private Const mstrDOCVAR_CONTACT As String = "Contact"
Private Const mstrDOCVAR_TITLE As String = "Title"
Private Const mstrDOCVAR_SCHOOL_DISTRICT As String = "SchoolDistrict"
Private Const mstrDOCVAR_ADDRESS As String = "Address"
Private Const mstrDOCVAR_PROGRAM As String = "Program"

'11/6/2013
Private Const mstrDOCVAR_SIGNATORY As String = "Signatory"

'10/17/2012
Private Const mstrDOCVAR_PROGRAM_SELECTION As String = "CurrentProgram"
Private Const mstrDOCVAR_CSPHONE As String = "CS Phone"

'6/13/2014
Private Const mstrFILE_NAME As String = "CBSoleSource"


'***********************************************************************************************************************************************
'11/13/2012
Private Const mstrBMK_LETTER_BODY As String = "LetterBody"

'4/24/2013
Private Const mstrBMK_COPIES As String = "NumCopies"

'11/13/2012
Private Const mstrDOCVAR_OFFICER As String = "SigningOfficer"

'***********************************************************************************************************************************************

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    'Re-architect to remove hard coded values
    'In fact consider storing in document variables that are NOT DISPLAYED (like I've always used)
    'Using document variables allows for users to change hard coded values themselves
    mstrContactInfoAutoText(ContactAutoText.ContactTitle) = "Cover Letter Contact"
    mstrContactInfoAutoText(ContactAutoText.ContactNoTitle) = "Cover Letter Contact - No Title"
    mstrContact = "CONTACT"
    mstrSchoolDistrict = "SCHOOL DISTRICT"
    mstrProgram = "PROGRAM"
    
    'Contract Specialist
    mstrContractSpecialistAutoText(CSAutoText.Nicole) = "CS Nicole"
    mstrContractSpecialistPhoneNum(CSAutoText.Nicole) = "212.373.8796"
    mstrContractSpecialistAutoText(CSAutoText.Adrian) = "CS Adrian"
    mstrContractSpecialistPhoneNum(CSAutoText.Adrian) = "212.373.8706"
    mstrContractSpecialistAutoText(CSAutoText.Alex) = "CS Alex"
    mstrContractSpecialistPhoneNum(CSAutoText.Alex) = "212.373.8787"
    
    
    'Current Region and Account Services Programs
    mstrProgramAutoText(ProgramAutoText.PN) = "PSAT"
    mstrProgramAutoText(ProgramAutoText.PN8_9) = "ReadiStep"
    mstrProgramAutoText(ProgramAutoText.PNRS) = "PSAT and ReadiStep"
    mstrProgramAutoText(ProgramAutoText.Path) = "Pathway"
    mstrProgramAutoText(ProgramAutoText.PD) = "Professional Development"
    mstrProgramAutoText(ProgramAutoText.AP) = "Advanced Placement"
    mstrProgramAutoText(ProgramAutoText.APPD) = "AP Professional"
    mstrProgramAutoText(ProgramAutoText.SAT) = "SAT"
    mstrProgramAutoText(ProgramAutoText.SB) = "SpringBoard"
    mstrProgramAutoText(ProgramAutoText.MS) = "Multiservice"
    mstrProgramAutoText(ProgramAutoText.Other) = "Program - Other"
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Public Property Let ContractSpecialist(ByRef renuCS As CSAutoText)
    On Error GoTo ErrorHandler
    
    mintActiveContractSpecialist = renuCS

'   This is more ledgeable code
'    If renuCS = Nicole Then
'        mintActiveContractSpecialist = Nicole
'    ElseIf renuCS = Peter Then
'        mintActiveContractSpecialist = Peter
'    End If
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Let Contact(ByRef rstrContact As String)
    On Error GoTo ErrorHandler
    
    If Len(rstrContact) <> 0 Then
        mstrContact = rstrContact
    End If
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Get Contact() As String
    On Error GoTo ErrorHandler
    
    If Len(mstrContact) <> 0 Then
        Contact = mstrContact
    End If
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Let Title(ByRef rstrTitle As String)
    On Error GoTo ErrorHandler
    
    If Len(rstrTitle) <> 0 Then
        mstrTitle = rstrTitle
        mintContactInfoChoice = ContactTitle
    Else
        mintContactInfoChoice = ContactNoTitle
    End If
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Get Title() As String
    On Error GoTo ErrorHandler
    
    If Len(mstrTitle) <> 0 Then
        Title = mstrTitle
        mintContactInfoChoice = ContactTitle
    Else
        mintContactInfoChoice = ContactNoTitle
    End If
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Get CSPhone() As String
    On Error GoTo ErrorHandler
    
    CSPhone = mstrContractSpecialistPhoneNum(mintActiveContractSpecialist)
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Let SchoolDistrict(ByRef rstrDistrict As String)
    On Error GoTo ErrorHandler
    
    If Len(rstrDistrict) <> 0 Then
        mstrSchoolDistrict = rstrDistrict
    End If
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Get SchoolDistrict() As String
    On Error GoTo ErrorHandler
    
    If Len(mstrSchoolDistrict) <> 0 Then
        SchoolDistrict = mstrSchoolDistrict
    End If
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Let Program(ByRef rstrProgram As String)
    On Error GoTo ErrorHandler
    
'    If Len(rstrProgram) <> 0 Then
'        mstrProgram = rstrProgram
'    End If
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Get Program() As String
    On Error GoTo ErrorHandler
    
'    If Len(mstrProgram) <> 0 Then
'        Program = mstrProgram
'    End If
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Let ContactAddress(ByRef rstrAddress As String)
    On Error GoTo ErrorHandler
    
    If Len(rstrAddress) <> 0 Then
        mstrAddress = Replace(rstrAddress, vbCrLf, vbVerticalTab)
    End If
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Let ProgramChoice(ByRef renuProgramChosen As ProgramAutoText)
    On Error GoTo ErrorHandler
    
    mintProgramChoice = renuProgramChosen
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Get ProgramChoice() As ProgramAutoText
    On Error GoTo ErrorHandler
    
    ProgramChoice = mintProgramChoice
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Get ContactAddress() As String
    On Error GoTo ErrorHandler
    
    If Len(mstrSchoolDistrict) <> 0 Then
        ContactAddress = mstrAddress
    End If
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Get ProgramList() As String()
    On Error GoTo ErrorHandler
    
    ProgramList = mstrProgramAutoText()
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Get ProgramListItem(ByRef rintIndex As Integer) As String
    On Error GoTo ErrorHandler
    
    ProgramListItem = mstrProgramAutoText(rintIndex)
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Sub ChangeContactInfo()
    On Error GoTo ErrorHandler
    
'    Application.Templates(ActiveDocument.AttachedTemplate.FullName).BuildingBlockEntries(mstrContactInfoAutoText(mintContactInfoChoice)).Insert _
'        Where:=ActiveDocument.Bookmarks(mstrBMK_CONTACT_INFO).Range, RichText:=True
        
    ActiveDocument.AttachedTemplate.BuildingBlockEntries(mstrContactInfoAutoText(mintContactInfoChoice)).Insert _
        Where:=ActiveDocument.Bookmarks(mstrBMK_CONTACT_INFO).Range, RichText:=True
        
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Public Sub InsertContractSpecialistAutoText()
    On Error GoTo ErrorHandler
    
    ActiveDocument.Bookmarks(mstrBMK_CLOSING).Select
'    Application.Templates(ActiveDocument.AttachedTemplate.FullName).BuildingBlockEntries(mstrContractSpecialistAutoText(mintActiveContractSpecialist)).Insert _
'        Where:=Selection.Range, RichText:=True
    
    ActiveDocument.AttachedTemplate.BuildingBlockEntries(mstrContractSpecialistAutoText(mintActiveContractSpecialist)).Insert _
        Where:=Selection.Range, RichText:=True
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Public Sub InsertProgram()
    On Error GoTo ErrorHandler
    
'    If AutoTextExist Then
'        ActiveDocument.Bookmarks(mstrBMK_PROGRAM).Select
'        Application.Templates(ThisDocument.FullName).BuildingBlockEntries(mstrProgramAutoText(mintProgramChoice)).Insert Where:=Selection.Range, RichText:=True
'        '10/17/2012 Must persist program choice
'        If DocVarExist(mstrDOCVAR_PROGRAM_SELECTION) Then
'            '10/17/2012 Created to store Program found in document or just selected by user
'            'Before user could not retrieve the program clearly on the cover letter into the form's drop down box
'            'This should work even for documents that don't have the document variable because it can be added.
'            ActiveDocument.Variables(mstrDOCVAR_PROGRAM_SELECTION).Value = mintProgramChoice
'        Else
'            ActiveDocument.Variables.Add Name:=mstrDOCVAR_PROGRAM_SELECTION, Value:=mintProgramChoice
'        End If
'    Else
'        MsgBox mstrProgramAutoText(mintProgramChoice) & " has not been created yet.  Please choose Program Other and enter your program manually", vbExclamation, "Insert Program"
'    End If
'
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Private Function AutoTextExist() As Boolean
    Dim objBldgBlockEntry As BuildingBlock
    Dim i As Integer
    Dim k As Integer
    On Error GoTo ErrorHandler
    
'    AutoTextExist = False
'    For i = 1 To Application.Templates(ThisDocument.FullName).BuildingBlockEntries.Count
'        Set objBldgBlockEntry = Application.Templates(ThisDocument.FullName).BuildingBlockEntries(i)
'        If objBldgBlockEntry.Name = mstrProgramAutoText(mintProgramChoice) Then
'            AutoTextExist = True
'            Exit For
'        End If
'    Next i
    
Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub CreateDocumentVariables()
    On Error GoTo ErrorHandler

    If Not DocVarExist(mstrDOCVAR_CONTACT) Then
        ActiveDocument.Variables.Add Name:=mstrDOCVAR_CONTACT, Value:=mstrDOCVAR_CONTACT
    End If
    
    If Not DocVarExist(mstrDOCVAR_TITLE) Then
        ActiveDocument.Variables.Add Name:=mstrDOCVAR_TITLE, Value:=mstrDOCVAR_TITLE
    End If
    
    If Not DocVarExist(mstrDOCVAR_SCHOOL_DISTRICT) Then
        ActiveDocument.Variables.Add Name:=mstrDOCVAR_SCHOOL_DISTRICT, Value:=mstrDOCVAR_SCHOOL_DISTRICT
    End If
    
    If Not DocVarExist(mstrDOCVAR_ADDRESS) Then
        ActiveDocument.Variables.Add Name:=mstrDOCVAR_ADDRESS, Value:=mstrDOCVAR_ADDRESS
    End If
    
    If Not DocVarExist(mstrDOCVAR_SIGNATORY) Then
        ActiveDocument.Variables.Add Name:=mstrDOCVAR_SIGNATORY, Value:=mstrDOCVAR_SIGNATORY
    End If
    
    If Not DocVarExist(mstrDOCVAR_PROGRAM) Then
        ActiveDocument.Variables.Add Name:=mstrDOCVAR_PROGRAM, Value:=mstrDOCVAR_PROGRAM
    End If
'
'    If Not DocVarExist(mstrDOCVAR_CSPHONE) Then
'        ActiveDocument.Variables.Add Name:=mstrDOCVAR_CSPHONE, Value:=mstrDOCVAR_CSPHONE
'    End If
'
    If Not DocVarExist(mstrDOCVAR_PROGRAM_SELECTION) Then
        '10/17/2012 Created to store Program found in document or just selected by user
        'Before user could not retrieve the program clearly on the cover letter into the form's drop down box
        'This should work even for documents that don't have the document variable because it can be added.
        ActiveDocument.Variables.Add Name:=mstrDOCVAR_PROGRAM_SELECTION, Value:=ProgramAutoText.PN
    End If
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Private Function DocVarExist(ByRef rstrDocVarName As String) As Boolean
    Dim objDocVar As Variable
    On Error GoTo ErrorHandler
    
    DocVarExist = False
    For Each objDocVar In ActiveDocument.Variables
        If objDocVar.Name = rstrDocVarName Then
            DocVarExist = True
            Debug.Print objDocVar.Name
            Exit For
        End If
    Next objDocVar
    
Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Function SaveDocument(ByRef rstrSchool As String) As String
    Const strPREFIX As String = "CB Sole Source - "

    Dim dlg As FileDialog
    On Error GoTo ErrorHandler
    
    Set dlg = Application.FileDialog(msoFileDialogSaveAs)
    'dlg.InitialFileName = mstrPATH_SOLE_SOURCE_LETTERS & strPREFIX & rstrSchool
    dlg.InitialFileName = mstrPATH_SOLE_SOURCE_LETTERS_v2 & strPREFIX & rstrSchool
    If dlg.Show Then
        SaveDocument = dlg.SelectedItems(1)
    End If
    
Exit_Here:
    Set dlg = Nothing
    Exit Function
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Property Get FileName() As String
    On Error GoTo ErrorHandler
    
    With ActiveDocument
        mstrFilePath = .CustomDocumentProperties("File Path").Value
        mstrFileName = .CustomDocumentProperties("File Name Prefix").Value & mstrSchoolDistrict & Space(1) & Format(Now(), "mmddyy")
        FileName = mstrFilePath & mstrFileName
    End With
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Let LetterType(ByRef renuLetterType As LetterType)
    On Error GoTo ErrorHandler
    
    mintLetterType = renuLetterType
    'Determine if this letter is a simple fully executed contract letter
    'or we are sending the customer 2 copies to be signed.
    ActiveDocument.Bookmarks(mstrBMK_LETTER_BODY).Select
    If LetterType = Signature Then
        ActiveDocument.AttachedTemplate.BuildingBlockEntries(mstrAUTOTEXT_SIGNATURE).Insert Where:=Selection.Range, RichText:=True
    ElseIf LetterType = FullyExecuted Then
        ActiveDocument.AttachedTemplate.BuildingBlockEntries(mstrAUTOTEXT_FULLY_EXECUTED).Insert Where:=Selection.Range, RichText:=True
    End If
        
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Get LetterType() As LetterType
    On Error GoTo ErrorHandler
    
    LetterType = mintLetterType
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property


Public Property Let NumOfCopies(ByRef rintCopyCount As Integer)
    Dim rng As Range
    Dim wrdPrecedeNumber As Range
    Dim rngWrd As Range
    Dim intWordIndex As Integer
    On Error GoTo ErrorHandler
    
    mintCopies = rintCopyCount
    If ActiveDocument.Bookmarks.Exists(mstrBMK_COPIES) Then
        Set rng = ActiveDocument.Bookmarks(mstrBMK_COPIES).Range
        rng.Select
        rng.Text = LCase(SpellNumber(mintCopies))
        rng.Select
        ActiveDocument.Bookmarks.Add Name:=mstrBMK_COPIES, Range:=rng
    End If

Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Get NumOfCopies() As Integer
    On Error GoTo ErrorHandler
    
    NumOfCopies = mintCopies
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Let Signatory(ByRef rstrSigningOfficer As String)
    On Error GoTo ErrorHandler
    
    If Len(rstrSigningOfficer) <> 0 Then
        mstrSignatory = rstrSigningOfficer
    Else
        mstrSignatory = "Vice President"
    End If
    ActiveDocument.Variables(mstrDOCVAR_OFFICER).Value = mstrSignatory
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

Public Property Get Signatory() As String
    On Error GoTo ErrorHandler
    
    If Len(mstrSignatory) <> 0 Then
        Signatory = mstrSignatory
    End If
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Property

