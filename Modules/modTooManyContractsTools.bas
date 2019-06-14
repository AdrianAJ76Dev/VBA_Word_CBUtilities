Attribute VB_Name = "modTooManyContractsTools"
Option Explicit

Public Sub MakeKhanEdit()
    Dim blnFound As Boolean
    Dim rngKhanTextTemp As Range
    Dim rngKhanText As Range
    
    On Error GoTo ErrorHandler
    
    Const mstrRIDER_KHAN_EDIT As String = "Client will receive access to comprehensive reporting"
    Const mintKHAN_EDIT_SENTENCES As String = 2
    
    'Text to edit
    With Selection.Find
        .Text = mstrRIDER_KHAN_EDIT
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        blnFound = .Execute
    End With
    If blnFound Then
        Set rngKhanTextTemp = Selection.Paragraphs(1).Range
        rngKhanTextTemp.Select
        Set rngKhanText = ActiveDocument.Range(Start:=rngKhanTextTemp.Sentences(3).Start, End:=rngKhanTextTemp.Sentences(4).End)
        rngKhanText.Select
    End If
    
Exit_Here:
    Set rngKhanText = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "MakeKhanEdit"
    End With
    Resume Exit_Here
End Sub

Public Sub RunFunction()
    Dim varRiders As Variant
    Dim fldRider As Field
    Dim idx As Integer
    
    Dim objContract As clsContract
    
    On Error GoTo ErrorHandler
    
    Set objContract = New clsCBContract
    varRiders = objContract.FindRidersUsingFieldObject
'    Debug.Print "Number of Riders in this document: " & UBound(varRiders)
    'Remember this contains the INDICES for the correct Fields that are RIDERS!
    For idx = 0 To UBound(varRiders)
        Set fldRider = varRiders(idx)
'        Debug.Print "Field # " & fldRider.Index
'        fldRider.Select
        fldRider.Unlink
    Next idx
    
    For Each fldRider In ActiveDocument.Fields
        If fldRider.Type = wdFieldIf Then
            fldRider.Code.Paragraphs(1).Range.Select
        End If
    Next fldRider

Exit_Here:
    Set objContract = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "RunFunction"
    End With
    Resume Exit_Here
End Sub

Public Sub TransferAutoText()
    
    Dim wrdSourceDoc As Document
    Dim wrdTargetDoc As Document
    Dim wrdTemplateSource As Template
    Dim wrdTemplateTarget As Template
    Dim bbEntries As BuildingBlockEntries
    Dim BBOld As BuildingBlock
    Dim BBNew As BuildingBlock
    Dim rng As Range
    
    Const strPATH As String = "C:\Documents and Settings\ajones\Application Data\Microsoft\Word\STARTUP\"
    Const strNAME_SOURCE As String = "CM Utilities v61.dotm"
    Const strNAME_TARGET As String = "CM Utilities v62.dotx"
    
    On Error GoTo ErrorHandler
    
    
'    Set wrdSourceDoc = Documents.Open(FileName:=strPATH & strNAME_SOURCE)
    Set wrdSourceDoc = ActiveDocument
    Set wrdTargetDoc = Documents.Open(FileName:=strPATH & strNAME_TARGET)
    
    Set wrdTemplateSource = wrdSourceDoc.AttachedTemplate
    Set wrdTemplateTarget = wrdTargetDoc.AttachedTemplate
    
    wrdTargetDoc.Activate
    Set bbEntries = wrdTemplateSource.BuildingBlockEntries
    For Each BBOld In bbEntries
        wrdTemplateTarget.BuildingBlockEntries.Add Name:=BBOld.Name, Type:=BBOld.Type, Category:=BBOld.Category, Range:=BBOld.Insert(Selection.Range, True)
        wrdTargetDoc.StoryRanges(wdMainTextStory).Delete
    Next BBOld
    
    
    

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "RunFunction"
    End With
    Resume Exit_Here
End Sub
