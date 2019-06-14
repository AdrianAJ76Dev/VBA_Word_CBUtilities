VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSoleSourceLetters 
   Caption         =   "Sole Source Letter"
   ClientHeight    =   7695
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   6840
   OleObjectBlob   =   "frmSoleSourceLetters.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSoleSourceLetters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mclsContent As clsCoverLetterVariables
Private mblnRetrieveVariables As Boolean
Private i As Integer

'Constants for AutoText Signatures
Private Const mstrSIGNATURE_JS As String = "Jeremy Singer"
Private Const mstrSIGNATURE_DM As String = "David C Meade Jr"
Private Const mstrSIGNATURE_TP As String = "Trevor Packer"

'Added additional signatories on 02/05/2016
Private Const mstrSIGNATURE_AC As String = "Auditi Chakravarty"

'Added additional signatories on 10/16/2017
Private Const mstrSIGNATURE_JD As String = "Jane Dapkus"

'Added signatory on 03/12/2019
Private Const mstrSIGNATURE_DW As String = "Douglas Waugh"


'Constants for AutoText Programs i.e. K12 Products vs. HED Products
Private Const mstrPRODUCTS_K12 As String = "SSL-K12"
Private Const mstrPRODUCTS_HED As String = "SSL-HED"
Private Const mstrPRODUCTS_PRICE_WARRANTY As String = "Sole Source - Price Warranty 2"

Private Const mstrHED_INSTITUTION As String = "Institution - College"
Private Const mstrK12_INSTITUTION As String = "Institution - School"
Private Const mstrPRICE_WARRANTY_INSTITUTION As String = "Institution Needing Price Warranty"

'10/3/2016
Private bmk As String

Private Const mstrAUTOTEXT_SSL As String = "Sole Source - Entire Letter Body"

Private Const mstrAUTOTEXT_PRICE_WARRANTY As String = "Sole Source - Price Warranty"
Private Const mstrAUTOTEXT_PRICE_WARRANTY_2 As String = "Sole Source - Price Warranty 2"
Private Const mstrAUTOTEXT_PRICE_WARRANTY_3 As String = "Sole Source - Price Warranty 3"

Private Const mstrBMK_SSL As String = "EntireLetterBody"
Private Const mstrBMK_PRICE_WARRANTY As String = "PriceWarranty"
Private Const mstrBMK_PRICE_WARRANTY_2 As String = "PriceWarranty2"
Private Const mstrBMK_PRICE_WARRANTY_3 As String = "PriceWarranty3"

Private mintPriceWarrantyNo As Integer

'Creating an enumeration for choice of Sole Source Letter, so I can include Price Warranty letters
Private Enum SSLType
    HED = 0
    K12 = 1
    PW = 2
End Enum

Private enuSelectedSSLType As SSLType

'1/27/2017
Private mobjSoleSourceLetter As Document

Private Function GetCurrentProgram() As String
    On Error GoTo ErrorHandler
    
'    GetCurrentProgram = ActiveDocument.Bookmarks("Program").Range.Text

Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Get Current Program"
    End With
    Resume Exit_Here
End Function

Private Sub cboProgram_Change()
    On Error GoTo ErrorHandler
    
'    mclsContent.ProgramChoice = Me.cboProgram.ListIndex
'    If mclsContent.ProgramChoice = Other Then
'        Me.txtProgram.Visible = True
'        Me.txtProgram.SetFocus
'    Else
'        Me.txtProgram.Visible = False
'    End If

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "cmdCancel_Click"
    End With
    Resume Exit_Here
End Sub


Private Sub cboPriceWarranty_Change()
    On Error GoTo ErrorHandler
    
    If Me.optPriceWarranty.Value = True Then
        mintPriceWarrantyNo = Val(Me.cboPriceWarranty.Value)
        ChooseSSLAutoText
    End If

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "cboPriceWarranty_Change"
    End With
    Resume Exit_Here
End Sub

Private Sub cboSignatory_Change()
    On Error GoTo ErrorHandler
    
    '01/31/2018 The line of code below is assuming AutoText SHOULD come from the Global Word Add-In
    'Application.Templates(ActiveDocument.AttachedTemplate.FullName).BuildingBlockEntries(Me.cboSignatory.Value).Insert _
        Where:=ActiveDocument.Bookmarks("Signature").Range, RichText:=True
        
    '01/31/2018 The line of code below is pulling the AutoText from the ATTACHED template
    ActiveDocument.AttachedTemplate.BuildingBlockEntries(Me.cboSignatory.Value).Insert _
        Where:=ActiveDocument.Bookmarks("Signature").Range, RichText:=True

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "cboSignatory_Change"
    End With
    Resume Exit_Here
End Sub

Private Sub cmdCancel_Click()
    Dim lngMsgResult As Long
    
    On Error GoTo ErrorHandler
    
    lngMsgResult = MsgBox("Close this document", vbYesNo + vbQuestion, "Create Contract Approval Letter")
    If lngMsgResult = vbYes Then
        ActiveDocument.Close SaveChanges:=False
    End If
    Unload Me

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "cmdCancel_Click"
    End With
    Resume Exit_Here
End Sub

Private Sub cmdClear_Click()
    On Error GoTo ErrorHandler
    
    Me.txtContact.Text = vbNullString
    Me.txtTitle.Text = vbNullString
    Me.txtAddress = vbNullString
    Me.txtSchoolDistrict.Text = vbNullString
'    Me.txtProgram.Text = vbNullString
    Me.RetrieveDocumentVariables = True
'    Me.cboProgram.ListIndex = 0
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "cmdClear_Click"
    End With
    Resume Exit_Here
End Sub

Private Sub cmdClose_Click()
    On Error GoTo ErrorHandler
    
    Unload Me

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "cmdClose_Click"
    End With
    Resume Exit_Here
End Sub

Private Sub cmdOk_Click()
    'This OK click button is for INSERTING new document variables AND RETRIEVING
    'New document variables.  When retrieving there's no need to do anything else
    On Error GoTo ErrorHandler
    
    If mblnRetrieveVariables Then
        '10/17/2012 Like the idea of folding the document variables into that class
        'The idea being the class models how the template should be constructed.
        'So Activedocument blah blah gets replaced by a class property
        Me.txtContact.Text = ActiveDocument.Variables("Contact").Value
        Me.txtTitle.Text = ActiveDocument.Variables("Title").Value
        Me.txtSchoolDistrict.Text = ActiveDocument.Variables("SchoolDistrict").Value
'        Me.txtProgram.Text = ActiveDocument.Variables("Program").Value
        Me.txtAddress.Text = ActiveDocument.Variables("Address").Value
'        Me.cboProgram.ListIndex = ActiveDocument.Variables("CurrentProgram").Value
    Else
        If Len(Me.txtContact.Text) <> 0 Then
            mclsContent.Contact = Me.txtContact.Text
            ActiveDocument.Variables("Contact").Value = mclsContent.Contact
        End If
        
        'Insert Contact AutoText with or without title based on whether
        'user has entered a title or not and switch to appropriate AutoText based on this
        'The class handles the AutoText assignment
        mclsContent.Title = Trim(Me.txtTitle.Text) 'TRIM in case user enters spaces and no text
        mclsContent.ChangeContactInfo
        If Len(Me.txtTitle.Text) <> 0 Then
            ActiveDocument.Variables("Title").Value = mclsContent.Title
        End If
        
        If Len(Me.txtSchoolDistrict.Text) <> 0 Then
            mclsContent.SchoolDistrict = Me.txtSchoolDistrict.Text
            ActiveDocument.Variables("SchoolDistrict").Value = mclsContent.SchoolDistrict
        End If
        
        'Either assign document variable OR insert appropriate AutoText
'        mclsContent.InsertProgram
'        If mclsContent.ProgramChoice = Other Then
'            Me.txtProgram.Enabled = True
'            If Len(Me.txtProgram.Text) <> 0 Then
'                mclsContent.Program = Me.txtProgram.Text
'                ActiveDocument.Variables("Program").Value = mclsContent.Program
'            End If
'        Else
'        End If
        
        If Len(Me.txtAddress.Text) <> 0 Then
            mclsContent.ContactAddress = Me.txtAddress.Text
            ActiveDocument.Variables("Address").Value = mclsContent.ContactAddress
        End If
        
'        mclsContent.ContractSpecialist = Me.cboContractSpecialist.ListIndex
'        mclsContent.InsertContractSpecialistAutoText
'        ActiveDocument.Variables("CS Phone").Value = mclsContent.CSPhone
        ActiveDocument.Fields.Update
    End If

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "cmdOK_Click"
    End With
    Resume Exit_Here
End Sub

Public Property Let RetrieveDocumentVariables(ByRef rblnRetrieveVariables As Boolean)
    On Error GoTo ErrorHandler
    
    If rblnRetrieveVariables Then
        Me.cmdOK.Accelerator = "R"
        Me.cmdOK.Caption = "Retrieve"
    Else
        Me.cmdOK.Accelerator = "O"
        Me.cmdOK.Caption = "OK"
    End If
    mblnRetrieveVariables = rblnRetrieveVariables
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Let RetrieveDocumentVariables"
    End With
    Resume Exit_Here
End Property

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler
    
     ActiveDocument.SaveAs2 FileName:=mclsContent.SaveDocument(Me.txtSchoolDistrict.Text)
   
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "txtContact_Change"
    End With
    Resume Exit_Here
End Sub

Private Sub optHED_Click()
    On Error GoTo ErrorHandler
    
    'Start up as HED i.e. Higher Education
    enuSelectedSSLType = HED
    Me.cboSignatory.ListIndex = 1 'David C. Meade Jr.
    Me.lblSchoolDistrict.Caption = mstrHED_INSTITUTION
    
    ChooseSSLAutoText
    
'    'Insert HED autotext now
'    Application.Templates(ThisDocument.FullName).BuildingBlockEntries(mstrPRODUCTS_HED).Insert _
'        Where:=ActiveDocument.Bookmarks(bmk).Range, RichText:=True
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "optHED_Click"
    End With
    Resume Exit_Here
End Sub

Private Sub optK12_Click()
    On Error GoTo ErrorHandler
    enuSelectedSSLType = K12
    Me.cboSignatory.ListIndex = 2 'Trevor Packer.
    Me.lblSchoolDistrict.Caption = mstrK12_INSTITUTION
    
    ChooseSSLAutoText
    
'    'Insert K12 autotext now
'    Application.Templates(ThisDocument.FullName).BuildingBlockEntries(mstrPRODUCTS_K12).Insert _
'        Where:=ActiveDocument.Bookmarks(bmk).Range, RichText:=True
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "optK12_Click"
    End With
    Resume Exit_Here
End Sub

Private Sub optPriceWarranty_Click()
    On Error GoTo ErrorHandler
    
    enuSelectedSSLType = PW
    mintPriceWarrantyNo = Val(Me.cboPriceWarranty.Value)
    Me.cboSignatory.ListIndex = 0 'Jeremy Singer.
    Me.lblSchoolDistrict.Caption = mstrPRICE_WARRANTY_INSTITUTION
    
    ChooseSSLAutoText
    
'    'Insert Price Warranty autotext now
'    Application.Templates(ThisDocument.FullName).BuildingBlockEntries(mstrPRODUCTS_PRICE_WARRANTY).Insert _
'        Where:=ActiveDocument.Bookmarks("EntireLetterBody").Range, RichText:=True
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "optPriceWarranty_Click"
    End With
    Resume Exit_Here
End Sub

Private Sub ChooseSSLAutoText()
    Dim bmkTemp As String
    Dim bmk As String
    Dim autoTextEntry As String
    
    On Error GoTo ErrorHandler
    
    'If Sole Source is currently a Price Warranty Letter
    'It's structure is quite different
    If ActiveDocument.Bookmarks.Exists(Name:=mstrBMK_PRICE_WARRANTY) _
        Or ActiveDocument.Bookmarks.Exists(Name:=mstrBMK_PRICE_WARRANTY_2) _
        Or ActiveDocument.Bookmarks.Exists(Name:=mstrBMK_PRICE_WARRANTY_3) _
    Then
        Select Case True
            Case ActiveDocument.Bookmarks.Exists(Name:=mstrBMK_PRICE_WARRANTY)
                bmkTemp = mstrBMK_PRICE_WARRANTY
                autoTextEntry = mstrAUTOTEXT_SSL
                
            Case ActiveDocument.Bookmarks.Exists(Name:=mstrBMK_PRICE_WARRANTY_2)
                bmkTemp = mstrBMK_PRICE_WARRANTY_2
                autoTextEntry = mstrAUTOTEXT_SSL
    
            Case ActiveDocument.Bookmarks.Exists(Name:=mstrBMK_PRICE_WARRANTY_3)
                bmkTemp = mstrBMK_PRICE_WARRANTY_3
                autoTextEntry = mstrAUTOTEXT_SSL
    
            Case Else
                bmkTemp = "Programs"
                
            
        End Select
    
        '01/31/2018 The line of code below is assuming AutoText SHOULD come from the Global Word Add-In
'        Application.Templates(ActiveDocument.AttachedTemplate.FullName).BuildingBlockEntries(autoTextEntry).Insert _
'            Where:=ActiveDocument.Bookmarks(bmkTemp).Range, RichText:=True
            
        '01/31/2018 The line of code below is pulling the AutoText from the ATTACHED template
        ActiveDocument.AttachedTemplate.BuildingBlockEntries(autoTextEntry).Insert _
            Where:=ActiveDocument.Bookmarks(bmkTemp).Range, RichText:=True
            
    End If
    
    bmkTemp = "Programs"
    Select Case enuSelectedSSLType
    
        Case SSLType.HED
            autoTextEntry = mstrPRODUCTS_HED
            
        Case SSLType.K12
            autoTextEntry = mstrPRODUCTS_K12
            
        Case SSLType.PW
            bmkTemp = mstrBMK_SSL
            If mintPriceWarrantyNo > 1 Then
                autoTextEntry = mstrAUTOTEXT_PRICE_WARRANTY & Space(1) & mintPriceWarrantyNo
            Else
                autoTextEntry = mstrAUTOTEXT_PRICE_WARRANTY
            End If
            
    End Select
    
    '01/31/2018 The line of code below is assuming AutoText SHOULD come from the Global Word Add-In
'    Application.Templates(ActiveDocument.AttachedTemplate.FullName).BuildingBlockEntries(autoTextEntry).Insert _
'        Where:=ActiveDocument.Bookmarks(bmkTemp).Range, RichText:=True
        
    '01/31/2018 The line of code below is pulling the AutoText from the ATTACHED template
    ActiveDocument.AttachedTemplate.BuildingBlockEntries(autoTextEntry).Insert _
        Where:=ActiveDocument.Bookmarks(bmkTemp).Range, RichText:=True
    
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "ChooseSSLAutoText"
    End With
    Resume Exit_Here
End Sub


Private Sub txtContact_Change()
    On Error GoTo ErrorHandler
    
    If Len(Me.txtContact.Text) <> 0 Then
        Me.RetrieveDocumentVariables = False
    End If
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "txtContact_Change"
    End With
    Resume Exit_Here
End Sub

Private Sub txtProgram_Change()
    On Error GoTo ErrorHandler
    
'    If Len(Me.txtProgram.Text) <> 0 Then
'        Me.RetrieveDocumentVariables = False
'    End If
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "txtProgram_Change"
    End With
    Resume Exit_Here
End Sub

Private Sub txtSchoolDistrict_Change()
    On Error GoTo ErrorHandler
    
    If Len(Me.txtSchoolDistrict.Text) <> 0 Then
        Me.RetrieveDocumentVariables = False
    End If
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "txtSchoolDistrict_Change"
    End With
    Resume Exit_Here
End Sub

Private Sub UserForm_Initialize()
    Dim intLower As Integer
    Dim intUpper As Integer
    Dim strCurrentItemOnList As String
    On Error GoTo ErrorHandler
    
    Set mclsContent = New clsCoverLetterVariables
    
    'Populate combo Box with Program Choices/AutoText Names
    Me.cboSignatory.AddItem mstrSIGNATURE_JS, 0
    Me.cboSignatory.AddItem mstrSIGNATURE_DM, 1
    Me.cboSignatory.AddItem mstrSIGNATURE_TP, 2
    
    'Added additional signatories on 02/05/2016
    Me.cboSignatory.AddItem mstrSIGNATURE_DW, 3
    Me.cboSignatory.AddItem mstrSIGNATURE_JD, 4
    Me.cboSignatory.AddItem mstrSIGNATURE_AC, 5
    
    
    Me.cboPriceWarranty.AddItem "1", 0
    Me.cboPriceWarranty.AddItem "2", 1
    Me.cboPriceWarranty.AddItem "3", 2
    
    'Added additional signatories on 02/05/2016
    Me.cboPriceWarranty.AddItem "4", 3
    Me.cboPriceWarranty.AddItem "5", 4
    Me.cboPriceWarranty.AddItem "6", 5
    
    
    'Start up as HED i.e. Higher Education
    Me.cboSignatory.ListIndex = 0 'Jeremy Singer
    Me.lblSchoolDistrict.Caption = mstrHED_INSTITUTION
       
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "UserForm_Initialize"
    End With
    Resume Exit_Here
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler
    
    If CloseMode = vbFormControlMenu Then
        
        Debug.Print "Cancel = " & Cancel
        Debug.Print "CloseMode = " & CloseMode
    
        ActiveDocument.Close SaveChanges:=False
    End If

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "UserForm_QueryClose"
    End With
    Resume Exit_Here
End Sub

Private Sub UserForm_Terminate()
    On Error GoTo ErrorHandler
    
    Set mclsContent = Nothing
    Set mobjSoleSourceLetter = Nothing
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "UserForm_Terminate"
    End With
    Resume Exit_Here
End Sub
