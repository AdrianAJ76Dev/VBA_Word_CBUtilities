Attribute VB_Name = "modContractTools"
Option Explicit

Private Const mstrNAME_CONTRACT_APPROVAL_LTR_BODY As String = "Contract Approval Letter"
Private Const mstrNAME_CONTRACT_APPROVAL_LTR_HEAD As String = "Contract Approval Letter Header"
Private Const mstrRIDER_HEADING_TEXT As String = "Schedule to College Board Enrollment Agreement"

Public Sub CreateContractApprovalLetter()
    On Error GoTo ErrorHandler
    
    'Determine if we are in a new fresh document
    If InStr(1, ActiveDocument.FullName, "Document") > 0 And _
        ActiveDocument.StoryRanges(wdMainTextStory).StoryLength = 1 Then
            Templates(ThisDocument.FullName).BuildingBlockEntries(mstrNAME_CONTRACT_APPROVAL_LTR_HEAD).Insert Where:=ActiveDocument.StoryRanges(wdPrimaryHeaderStory)
            Templates(ThisDocument.FullName).BuildingBlockEntries(mstrNAME_CONTRACT_APPROVAL_LTR_BODY).Insert Where:=Selection.Range
    Else
        MsgBox "This document appears not to be completely blank.  Please close this document and create a new blank document.", vbInformation, "Create Contract Approval Letter"
    End If

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Create Contract Approval Letter"
    End With
    Resume Exit_Here
End Sub

Public Sub FormatProcessContract()
    On Error GoTo ErrorHandler
    
    FormatContract
    ProcessContract

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Format and Process the Contract"
    End With
    Resume Exit_Here
End Sub

Public Sub FormatContract()
    Dim rngContract As Range
    On Error GoTo ErrorHandler
    
    '03/08/2013 - Common Contract Specialist practice is to remove
    'all space before and after in the document.
    ActiveDocument.TrackRevisions = True
    ActiveDocument.TrackFormatting = True

    ActiveDocument.StoryRanges(wdMainTextStory).Select
    Set rngContract = Selection.Range
    rngContract.ParagraphFormat.SpaceBefore = 0
    rngContract.ParagraphFormat.SpaceAfter = 0
    rngContract.Font.Name = "Times New Roman"
    rngContract.Font.Size = 10
    '
    '
    ' Replace2012 Macro
    ' Search and Replace 2012-2013 with 2013-2014
    '
    rngContract.Find.ClearFormatting
    rngContract.Find.Replacement.ClearFormatting
    With rngContract.Find
        .Text = "2012-2013"
        .Replacement.Text = "2013-2014"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    rngContract.Find.Execute Replace:=wdReplaceAll
    rngContract.Collapse Direction:=wdCollapseStart
    ActiveDocument.Fields.Update

Exit_Here:
    Set rngContract = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Format Contract"
    End With
    Resume Exit_Here
End Sub

Public Sub ProcessContract()
    Const strVP_FIRST_NAME As String = "Stacy"
    Const strVP_LAST_NAME As String = "Caldwell"
    Const strVP_TITLE As String = "VP District & State Assessment Programs"
    
    Const strSVP_FIRST_NAME As String = "Tom"
    Const strSVP_LAST_NAME As String = "Higgins"
    Const strSVP_TITLE As String = "SVP, AP and Instruction"
    
    Const strDP_TOTAL As String = "Contract Total"
    Const strDP_FNAME As String = "CB First Name"
    Const strDP_LNAME As String = "CB Last Name"
    Const strDP_TITLE As String = "CB Job Title"
    
    Const strDP_REVENUE_NO As String = "Contract Number"
    Const strDP_DISTRICT As String = "Company Name"
    Const strFILENAME_PSAT As String = " PN EPP "
    Const strFILENAME_FY As String = "2013-2014 "
    
    Const strDP_START_DATE As String = "Contract Begin Date"
    Const strDP_CONTRACT_DAY As String = "Day of Contract Begin Date"
    Const strDP_CONTRACT_YEAR As String = "Year of Contract Begin Date"
    Const strDP_CONTRACT_MONTH As String = "Month of Contract Begin Date"
    
'    Const strCONTRACT_START_DAY As String = "April 1, 2013"
    
    Dim sngQuote As Single
    Dim strFileName As String
    Dim tbl As Table
    
    Const strPATH_CONTRACTS As String = "Q:\RAS Contracts Management\Pivotal\K12 Contracts - Pivotal\"
    Dim fd As FileDialog
    
    '05/17/2013
    Dim astr
    
    On Error GoTo ErrorHandler
    
    'Save the document
    If ActiveDocument.Saved = False Then
        strFileName = ActiveDocument.CustomDocumentProperties(strDP_DISTRICT)
        strFileName = strFileName & strFILENAME_PSAT
        strFileName = strFileName & strFILENAME_FY
        strFileName = strFileName & ActiveDocument.CustomDocumentProperties(strDP_REVENUE_NO)
        strFileName = strFileName & " rl"
        
        Set fd = Application.FileDialog(msoFileDialogSaveAs)
        With fd
            .InitialFileName = strPATH_CONTRACTS & strFileName
            If .Show = 0 Then 'user pressed the cancel button
            Else
                .Execute
            End If
        End With
        Set fd = Nothing
    End If
    
    With ActiveDocument
        'Populate the Senior Vice/Vice President
        sngQuote = .CustomDocumentProperties(strDP_TOTAL).Value
        If sngQuote < 100000 Then
            .CustomDocumentProperties(strDP_FNAME).Value = strVP_FIRST_NAME
            .CustomDocumentProperties(strDP_LNAME).Value = strVP_LAST_NAME
            .CustomDocumentProperties(strDP_TITLE).Value = strVP_TITLE
        Else
            .CustomDocumentProperties(strDP_FNAME).Value = strSVP_FIRST_NAME
            .CustomDocumentProperties(strDP_LNAME).Value = strSVP_LAST_NAME
            .CustomDocumentProperties(strDP_TITLE).Value = strSVP_TITLE
        End If
        .CustomDocumentProperties(strDP_START_DATE).Value = Format(DateAdd("m", 1, Now()), "mmmm 1, yyyy")
        .CustomDocumentProperties(strDP_CONTRACT_MONTH).Value = Format(Now(), "mmmm")
        
        'IIf(Right(Format(Now(), "d"), 1) > 3, "th", Choose(Right(Format(Now(), "d"), 1), "st", "nd", "rd"))
        'The logic here is to look at the 1st digit of the current day.  For example 17 you would look at 7
        'Since 1, 2, 3 are the only values that have differing suffixes, you focus on them
        'Therefore Choose(1, "st", "nd", "rd")) returns 1st
        'Every digit from 4 - 9 the suffix is "th"
        .CustomDocumentProperties(strDP_CONTRACT_DAY).Value = Format(Now(), "d") _
            & IIf(Format(Now(), "d") > 3, "th", Choose(Right(Format(Now(), "d"), 1), "st", "nd", "rd", "th", "th", "th", "th", "th", "th", "th", "th", "th", "th"))
        .CustomDocumentProperties(strDP_CONTRACT_YEAR).Value = Format(Now(), "yyyy")
        .Fields.Update
        
        For Each tbl In .Tables
            tbl.Rows.WrapAroundText = False
        Next tbl
    End With
    

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Process Contract"
    End With
    Resume Exit_Here
End Sub

Public Sub CreateHEDAgreement()
    Dim rngContract As Range
    Const strDOCVAR_CLIENT As String = "Short College Name"
    
    '8/18/2014 Update Hyperlinks for "Uses of College Board Test Scores & Related Data"
    Const strHYPERLINK_TEST_SCORES_RELATED_DATA As String = "http://www.collegeboard.com/research/home"
    
    On Error GoTo ErrorHandler

    'Remove Space After, format to standard HED contract font: Times New Roman 11 pt. and turn track changes on
    ActiveDocument.TrackRevisions = True
    ActiveDocument.TrackFormatting = True

    Set rngContract = ActiveDocument.StoryRanges(wdMainTextStory)
    rngContract.ParagraphFormat.SpaceBefore = 0
    rngContract.ParagraphFormat.SpaceAfter = 0
    rngContract.Font.Name = "Times New Roman"
    rngContract.Font.Size = 11
    
    ActiveDocument.CustomDocumentProperties.Add _
                    Name:=strDOCVAR_CLIENT, LinkToContent:=False, Value:="Client", _
                    Type:=msoPropertyTypeString
                    
    ActiveDocument.Fields.Update
    
    'Locate the Riders
    With rngContract.Find
        .Text = mstrRIDER_HEADING_TEXT
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    If rngContract.Find.Execute = True Then
        rngContract.Select
    End If
    
    '8/18/2014 Update Hyperlinks for "Uses of College Board Test Scores & Related Data"
    With rngContract.Find
        .Text = strHYPERLINK_TEST_SCORES_RELATED_DATA
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    If rngContract.Find.Execute = True Then
        rngContract.Select
        
    End If
    
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "CreateHEDAgreement"
    End With
    Resume Exit_Here
End Sub

