Attribute VB_Name = "modCoverLetterMain"
Option Explicit
Dim mclsContent As clsCoverLetterVariables
Private Const mstrDOCVAR_PROGRAM As String = "Program"
Private mobjCoverLetter As Document

Public Sub CreateCoverLetter()
    On Error GoTo ErrorHandler
    
    Set mobjCoverLetter = Documents.Add(Template:=CoverLetterTemplateFullName)
    mobjCoverLetter.Activate
    
    frmCoverLetter.RetrieveDocumentVariables = True
    frmCoverLetter.Show
     
Exit_Here:
    Set mobjCoverLetter = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "CreateCoverLetter"
    End With
    Resume Exit_Here
End Sub

Public Sub RetrieveDocumentVariables()
    On Error GoTo ErrorHandler
    
    frmCoverLetter.Show
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "RetrieveDocumentVariables"
    End With
    Resume Exit_Here
End Sub

Public Sub CreateDocumentVariables()
    On Error GoTo ErrorHandler
    
    Set mclsContent = New clsCoverLetterVariables
    mclsContent.CreateDocumentVariables
    
Exit_Here:
    Set mclsContent = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Create Document Variables"
    End With
    Resume Exit_Here
End Sub

Public Sub DocVarExist()
    Dim objDocVar As Variable
    On Error GoTo ErrorHandler
    
    For Each objDocVar In ActiveDocument.Variables
        If objDocVar.Name = mstrDOCVAR_PROGRAM Then
            Debug.Print "THIS IS THE ONE I WANT ----> " & objDocVar.Name
        Else
            Debug.Print objDocVar.Name
        End If
    Next objDocVar

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Public Sub ShowDocumentVariables()
    On Error GoTo ErrorHandler
    
    frmDocumentVariables.Show

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "ShowDocumentVariables"
    End With
    Resume Exit_Here
End Sub

Private Function CoverLetterTemplateFullName() As String
    Dim fso As FileSystemObject
    Const mstrPATH_USER_TEMPLATES As String = "\Microsoft\Templates\Contracts Management\"
    
'    Const mstrPATH_TEMPLATES As String = "C:\Users\ajones\AppData\Roaming\Microsoft\Templates\CM Current Templates\"
    '5/5/2017 Use the constant below instead of the hardcoded bath since it's SPECIFIC to me an my machine.
    Const USERPROFILE As String = "USERPROFILE"
    Const USR_APPDATA As String = "APPDATA"
    Const mstrTEMPLATE_SOLE_SOURCE As String = "Send Cover Letter for FullyExecuted.dotx"

    On Error GoTo ErrorHandler
    Set fso = New FileSystemObject
    
    CoverLetterTemplateFullName = Environ$(USR_APPDATA) & mstrPATH_USER_TEMPLATES & mstrTEMPLATE_SOLE_SOURCE
    
    If Not fso.FileExists(CoverLetterTemplateFullName) Then
        Err.Raise Number:=vbObjectError + 1, Source:="", _
            Description:="File Name " & mstrTEMPLATE_SOLE_SOURCE & " could not be found in:" & vbCr _
            & mstrPATH_USER_TEMPLATES
    End If

Exit_Here:
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "CoverLetterTemplateFullName"
    End With
    Resume Exit_Here
End Function



