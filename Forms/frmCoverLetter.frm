VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCoverLetter 
   Caption         =   "Cover Letter"
   ClientHeight    =   6225
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9555
   OleObjectBlob   =   "frmCoverLetter.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCoverLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mclsContent As clsCoverLetterVariables
Private mblnRetrieveVariables As Boolean
Private i As Integer

Private Function GetCurrentProgram() As String
    On Error GoTo ErrorHandler
    
    GetCurrentProgram = ActiveDocument.Bookmarks("Program").Range.Text

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
    
    mclsContent.ProgramChoice = Me.cboProgram.ListIndex
    ActiveDocument.Variables("Program").Value = Me.cboProgram.Value
    ActiveDocument.Fields.Update
    If mclsContent.ProgramChoice = Other Then
        Me.txtProgram.Visible = True
        Me.txtProgram.SetFocus
    Else
        Me.txtProgram.Visible = False
    End If

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "cmdCancel_Click"
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
    Me.txtProgram.Text = vbNullString
    Me.RetrieveDocumentVariables = True
    Me.cboProgram.ListIndex = 0
    
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
    
    Application.FileDialog(msoFileDialogSaveAs).InitialFileName = mclsContent.FileName
    Application.FileDialog(msoFileDialogSaveAs).Show
    Application.FileDialog(msoFileDialogSaveAs).Execute
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
    
    UpdateCoverLetter

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
        Me.cmdOK.Accelerator = "U"
        Me.cmdOK.Caption = "Update"
    End If
    mblnRetrieveVariables = rblnRetrieveVariables
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "UserForm_Initialize"
    End With
    Resume Exit_Here
End Property

Private Sub optCoverLetterFullyExecuted_Click()
    On Error GoTo ErrorHandler
    
    Me.lblNumOfCopies.Enabled = False
    Me.txtNumOfCopies.Enabled = False
    mclsContent.LetterType = FullyExecuted
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "optCoverLetterFullyExecuted_Click"
    End With
    Resume Exit_Here
End Sub

Private Sub optCoverLetterSignature_Click()
    On Error GoTo ErrorHandler
    
    Me.lblNumOfCopies.Enabled = True
    Me.txtNumOfCopies.Enabled = True
    mclsContent.LetterType = Signature
    mclsContent.NumOfCopies = txtNumOfCopies.Text
    mclsContent.Signatory = Me.txtSignatory

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "optCoverLetterSignature_Click"
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

Private Sub txtNumOfCopies_AfterUpdate()
    On Error GoTo ErrorHandler
    
    If Len(Me.txtNumOfCopies.Text) <> 0 Then
        Me.RetrieveDocumentVariables = False
    End If
    mclsContent.NumOfCopies = txtNumOfCopies.Text
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "txtNumOfCopies_AfterUpdate"
    End With
    Resume Exit_Here
End Sub

Private Sub txtProgram_Change()
    On Error GoTo ErrorHandler
    
    If Len(Me.txtProgram.Text) <> 0 Then
        Me.RetrieveDocumentVariables = False
    End If
    
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

Private Sub txtSignatory_AfterUpdate()
'    On Error GoTo ErrorHandler
'
'    mclsContent.Signatory = Me.txtSignatory
'    ActiveDocument.Fields.Update
'
'Exit_Here:
'    Exit Sub
'
'ErrorHandler:
'    With Err
'        MsgBox .Number & vbCr & .Description, vbCritical, "txtSignatory_AfterUpdate"
'    End With
'    Resume Exit_Here
End Sub

Private Sub txtSignatory_Change()
    On Error GoTo ErrorHandler
    
    mclsContent.Signatory = Me.txtSignatory
    ActiveDocument.Fields.Update
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "txtSignatory_Change"
    End With
    Resume Exit_Here
End Sub

Private Sub UserForm_Initialize()
    Dim intLower As Integer
    Dim intUpper As Integer
    Dim strCurrentItemOnList As String
    On Error GoTo ErrorHandler
    
    Set mclsContent = New clsCoverLetterVariables
    Me.cboContractSpecialist.AddItem "Nicole McIntyre", 0
    Me.cboContractSpecialist.AddItem "Adrian Jones", 1
    Me.cboContractSpecialist.AddItem "Alexandra Stabilito", 2
    Me.cboContractSpecialist.ListIndex = 0
    Me.txtProgram.Enabled = True
    Me.txtProgram.Visible = False
    
    'Populate combo Box with Program Choices/AutoText Names
    intLower = LBound(mclsContent.ProgramList())
    intUpper = UBound(mclsContent.ProgramList())
    For i = intLower To intUpper
        strCurrentItemOnList = mclsContent.ProgramListItem(i)
        Me.cboProgram.AddItem strCurrentItemOnList, i
    Next
    Me.cboProgram.ListIndex = intLower
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "UserForm_Initialize"
    End With
    Resume Exit_Here
End Sub

Private Sub UserForm_Terminate()
    On Error GoTo ErrorHandler
    
    Set mclsContent = Nothing
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "UserForm_Terminate"
    End With
    Resume Exit_Here
End Sub

Private Sub ChooseLetterType()
    On Error GoTo ErrorHandler

    MsgBox "Clicked Cover Letter Type Frame", vbInformation, "fraCoverLetterType_Click"
'    If Me.optCoverLetterSignature Then
'        mclsContent.LetterType = Signature
'    Else
'        mclsContent.LetterType = FullyExecuted
'    End If
        
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "fraCoverLetterType_Click"
    End With
    Resume Exit_Here
End Sub

Private Sub UpdateCoverLetter()
    On Error GoTo ErrorHandler
    
    If mblnRetrieveVariables Then
        '10/17/2012 Like the idea of folding the document variables into that class
        'The idea being the class models how the template should be constructed.
        'So Activedocument blah blah gets replaced by a class property
        Me.txtContact.Text = ActiveDocument.Variables("Contact").Value
        Me.txtTitle.Text = ActiveDocument.Variables("Title").Value
        Me.txtSchoolDistrict.Text = ActiveDocument.Variables("SchoolDistrict").Value
        Me.txtProgram.Text = ActiveDocument.Variables("Program").Value
        Me.txtAddress.Text = ActiveDocument.Variables("Address").Value
        Me.cboProgram.ListIndex = ActiveDocument.Variables("CurrentProgram").Value
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
        mclsContent.InsertProgram
        If mclsContent.ProgramChoice = Other Then
            Me.txtProgram.Enabled = True
            If Len(Me.txtProgram.Text) <> 0 Then
                mclsContent.Program = Me.txtProgram.Text
                ActiveDocument.Variables("Program").Value = mclsContent.Program
            End If
        Else
        End If
        
        If Len(Me.txtAddress.Text) <> 0 Then
            mclsContent.ContactAddress = Me.txtAddress.Text
            ActiveDocument.Variables("Address").Value = mclsContent.ContactAddress
        End If
        
        mclsContent.ContractSpecialist = Me.cboContractSpecialist.ListIndex
        mclsContent.InsertContractSpecialistAutoText
        ActiveDocument.Variables("CS Phone").Value = mclsContent.CSPhone
        ActiveDocument.Fields.Update
    End If
        
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Update Cover Letter"
    End With
    Resume Exit_Here
End Sub
