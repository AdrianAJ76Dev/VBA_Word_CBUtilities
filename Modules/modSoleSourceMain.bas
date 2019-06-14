Attribute VB_Name = "modSoleSourceMain"
Option Explicit

Dim mclsContent As clsCoverLetterVariables
Private Const mstrDOCVAR_PROGRAM As String = "Program"
'1/30/2017
Private Const mstrTEMPLATE_SOLE_SOURCE As String = "Sole Source Letter v2.dotm"
Private wrdDoc As Document
Private mobjSoleSourceLetter As Document

Public Sub CreateSoleSourceLetter()
    Const mstrTEMPLATE_SOLE_SOURCE As String = "Sole Source Letter v5.dotx"
    On Error GoTo ErrorHandler
    
    Set mobjSoleSourceLetter = Documents.Add(PullFromSharePoint(mstrTEMPLATE_SOLE_SOURCE))
    mobjSoleSourceLetter.Activate

    frmSoleSourceLetters.RetrieveDocumentVariables = False
    frmSoleSourceLetters.Show
     
Exit_Here:
    Set mobjSoleSourceLetter = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "CreateSoleSourceLetter"
    End With
    Resume Exit_Here
End Sub

Private Function SoleSourceTemplateFullName() As String
    '01/31/2018 Retrieve the template from SharePoint
    Const mstrPATH_SHAREPOINT As String = "https://mysharepoint/sdp/oosvp/RAS/Contracts Management/"
    Const mstrPATH_OFFICE_TEMPLATE As String = "C:\Documents and Settings\ajones\Application Data\Microsoft\Templates\Contracts Management\Automated Templates\"
    Const mstrTEMPLATE_SOLE_SOURCE As String = "Sole Source Letter v5.dotx"

    SoleSourceTemplateFullName = mstrPATH_SHAREPOINT & mstrTEMPLATE_SOLE_SOURCE
'    SoleSourceTemplateFullName = mstrPATH_OFFICE_TEMPLATE & mstrTEMPLATE_SOLE_SOURCE


Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "SoleSourceTemplateFullName"
    End With
    Resume Exit_Here
End Function

Public Sub EditSoleSourceLetter()
    On Error GoTo ErrorHandler
    
    frmSoleSourceLetters.RetrieveDocumentVariables = True
    frmSoleSourceLetters.Show
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "EditSoleSourceLetter"
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
        MsgBox .Number & vbCr & .Description, vbCritical, "RetrieveDocumentVariables"
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

Private Function SoleSourceTemplateFullNameOLD() As String
    Dim fso As FileSystemObject
    Const mstrPATH_USER_TEMPLATES As String = "\Microsoft\Templates\Contracts Management\"
    
    Dim xmlhttp As MSXML2.XMLHTTP60
    Const mstrPATH_SHAREPOINT As String = "https://mysharepoint/sdp/oosvp/RAS/Contracts Management/"
    
'    Const mstrPATH_TEMPLATES As String = "C:\Users\ajones\AppData\Roaming\Microsoft\Templates\CM Current Templates\"
    '5/5/2017 Use the constant below instead of the hardcoded bath since it's SPECIFIC to me an my machine.
    
    Const USERPROFILE As String = "USERPROFILE"
    Const USR_APPDATA As String = "APPDATA"
    Const mstrTEMPLATE_SOLE_SOURCE As String = "Sole Source Letter v5.dotx"

    On Error GoTo ErrorHandler
    Set fso = New FileSystemObject
    
    'SoleSourceTemplateFullName = Environ$(USR_APPDATA) & mstrPATH_USER_TEMPLATES & mstrTEMPLATE_SOLE_SOURCE
    SoleSourceTemplateFullNameOLD = mstrPATH_SHAREPOINT & mstrTEMPLATE_SOLE_SOURCE
    
'    If Not fso.FileExists(SoleSourceTemplateFullName) Then
'        Err.Raise Number:=vbObjectError + 1, Source:="", _
'            Description:="File Name " & mstrTEMPLATE_SOLE_SOURCE & " could not be found in:" & vbCr _
'            & mstrPATH_USER_TEMPLATES
'    End If

    
'    If Not fso.FileExists(SoleSourceTemplateFullName) Then
'        Err.Raise Number:=vbObjectError + 1, Source:="", _
'            Description:="File Name " & mstrTEMPLATE_SOLE_SOURCE & " could not be found in:" & vbCr _
'            & mstrPATH_SHAREPOINT
'    End If


Exit_Here:
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "SoleSourceTemplateFullNameOLD"
    End With
    Resume Exit_Here
End Function
'
'Public Sub ConnectToSharePointUsingServer()
'    Dim svrxmlhttp As MSXML2.ServerXMLHTTP60
'    Const mstrTEMPLATE_SOLE_SOURCE As String = "Sole Source Letter v3.dotx"
'    Const mstrPATH_SHAREPOINT As String = "https://mysharepoint/sdp/oosvp/RAS/Contracts Management/"
'
'    On Error GoTo ErrorHandler
'
'    Set svrxmlhttp = New MSXML2.ServerXMLHTTP60
'    svrxmlhttp.Open "GET", mstrPATH_SHAREPOINT + mstrTEMPLATE_SOLE_SOURCE, False
'    svrxmlhttp.send ("Username=ajones@collegeboard.org&Password=WimWenders16")
'    MsgBox svrxmlhttp.Status & vbCr & svrxmlhttp.StatusText
'
'
'Exit_Here:
'    Exit Sub
'
'ErrorHandler:
'    With Err
'        MsgBox .Number & vbCr & .Description, vbCritical, "ConnectToSharePoint"
'    End With
'    Resume Exit_Here
'End Sub

Public Sub ConnectToSharePoint()
    Dim xmlhttp As MSXML2.XMLHTTP60
    Const mstrTEMPLATE_SOLE_SOURCE As String = "Sole Source Letter v5.dotx"
    Const mstrPATH_SHAREPOINT As String = "https://mysharepoint/sdp/oosvp/RAS/Contracts Management/"
    
    On Error GoTo ErrorHandler

    Set xmlhttp = New MSXML2.ServerXMLHTTP60
    xmlhttp.Open "GET", mstrPATH_SHAREPOINT + mstrTEMPLATE_SOLE_SOURCE, False
    xmlhttp.send
    MsgBox xmlhttp.Status & vbCr & xmlhttp.StatusText
    
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "ConnectToSharePoint"
    End With
    Resume Exit_Here
End Sub


