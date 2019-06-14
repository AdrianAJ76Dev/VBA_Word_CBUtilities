Attribute VB_Name = "HEDTools"
Option Explicit

Private Const mstrRIDER_HEADING_TEXT As String = "Schedule to College Board Enrollment Agreement"

Public Sub CreateHEDAgreement()
    Dim rngContract As Range
    Dim rngAutoTextGoesHere As Range
    Dim blnFound As Boolean
    
    Const strDOCVAR_CLIENT As String = "Short College Name"
    
    '8/18/2014 Update Hyperlinks for "Uses of College Board Test Scores & Related Data"
    Const strHYPERLINK_TEST_SCORES_RELATED_DATA As String = "http://www.collegeboard.com/research/home"
    Const strAUTOTEXT_HYPERLINK As String = "HGH - HED Links - Guidelines"
    
    'Added 12/08/2015 - The Hyperlinks below are the correct Hyperlinks and they work unlike the originals
'    Const strHYPERLINK_TEST_SCORES_RELATED_DATA As String = "http://www.collegeboard.com/research/home"
'    Const strAUTOTEXT_HYPERLINK As String = "HGH - HED Links - Guidelines"
'
'    Const strHYPERLINK_TEST_SCORES_RELATED_DATA As String = "http://www.collegeboard.com/research/home"
'    Const strAUTOTEXT_HYPERLINK As String = "HGH - HED Links - Guidelines"
    
    Const intTEMPLATE_BB As Integer = 4 'Building Blocks.dotx --- where common autotext is stored
    
    Dim objCBUtilities As Template
    
    '8/19/2014
    Dim objDocProp As DocumentProperty
    
    On Error GoTo ErrorHandler

    'Remove Space After, format to standard HED contract font: Times New Roman 11 pt. and turn track changes on
    ActiveDocument.TrackRevisions = True
    ActiveDocument.TrackFormatting = True
    Templates.LoadBuildingBlocks
    Set objCBUtilities = Templates(intTEMPLATE_BB)
    
    Set rngContract = ActiveDocument.StoryRanges(wdMainTextStory)
    rngContract.ParagraphFormat.SpaceBefore = 0
    rngContract.ParagraphFormat.SpaceAfter = 0
    rngContract.Font.Name = "Times New Roman"
    rngContract.Font.Size = 11
    
    For Each objDocProp In ActiveDocument.CustomDocumentProperties
        If objDocProp.Name = strDOCVAR_CLIENT Then
            blnFound = True
        End If
    Next objDocProp
    
    If blnFound = False Then
        ActiveDocument.CustomDocumentProperties.Add _
            Name:=strDOCVAR_CLIENT, LinkToContent:=False, Value:="Client", _
            Type:=msoPropertyTypeString
    End If
                    
    ActiveDocument.Fields.Update
    
    InsertPricingTableAndPaymentSchedule
    
    '8/18/2014 Update Hyperlinks for "Uses of College Board Test Scores & Related Data"
    'Add Do While
    ActiveDocument.TrackRevisions = False
    ActiveDocument.TrackFormatting = False
    
    Do
        With rngContract.Find
            .Text = strHYPERLINK_TEST_SCORES_RELATED_DATA
            .Forward = True
            .Wrap = wdFindAsk
            .Format = True
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            blnFound = .Execute(Forward:=True)
        End With
        
        If blnFound Then
            rngContract.Select 'This is how I select what was found
            Set rngAutoTextGoesHere = Selection.Range
            objCBUtilities.BuildingBlockEntries(strAUTOTEXT_HYPERLINK).Insert Where:=rngAutoTextGoesHere
            Selection.Collapse Direction:=wdCollapseEnd
        End If
    Loop Until Not blnFound
    ActiveDocument.TrackRevisions = True
    ActiveDocument.TrackFormatting = True
        
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
        blnFound = .Execute
    End With
    If rngContract.Find.Execute = True Then
        rngContract.Select
    End If

Exit_Here:
    Set rngContract = Nothing
    Set rngAutoTextGoesHere = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "CreateHEDAgreement"
    End With
    Resume Exit_Here
End Sub

Public Sub InsertPricingTableAndPaymentSchedule()
    Dim rngHEDAgreement As Range
    Dim rngAutoTextGoesHere As Range
    Dim blnFound As Boolean
    Dim objCBUtilities As Template
    
    Const strPRICING As String = "[insert pricing table]"
    Const strAT_PRICING_TABLE As String = "Pricing Table"
    Const strPAYMENT As String = "[insert payment schedule]"
    Const strAT_PAYMENT As String = "Payment Schedule"
    Const intTEMPLATE_BB As Integer = 3 'Building Blocks.dotx --- where common autotext is stored
    
    'AutoText for Pricing Table and Payment Schedule
    On Error GoTo ErrorHandler
    
    ActiveDocument.TrackRevisions = False
    ActiveDocument.TrackFormatting = False
    blnFound = False
    Templates.LoadBuildingBlocks
    Set objCBUtilities = Templates(intTEMPLATE_BB)
    Set rngHEDAgreement = ActiveDocument.StoryRanges(wdMainTextStory)
    rngHEDAgreement.Find.ClearFormatting
    rngHEDAgreement.Find.Replacement.ClearFormatting
    Do
        With rngHEDAgreement.Find
            .Text = strPRICING
            .Forward = True
            .Wrap = wdFindAsk
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            blnFound = .Execute(Forward:=True)
        End With
        
        If blnFound Then
            rngHEDAgreement.Select
            Set rngAutoTextGoesHere = Selection.Range
            objCBUtilities.BuildingBlockEntries(strAT_PRICING_TABLE).Insert Where:=rngAutoTextGoesHere
        End If
        
        With rngHEDAgreement.Find
            .Text = strPAYMENT
            .Forward = True
            .Wrap = wdFindAsk
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            blnFound = .Execute(Forward:=True)
        End With
        
        If blnFound Then
            rngHEDAgreement.Select
            Set rngAutoTextGoesHere = Selection.Range
            objCBUtilities.BuildingBlockEntries(strAT_PAYMENT).Insert Where:=rngAutoTextGoesHere
        End If
Loop Until Not blnFound

Exit_Here:
    ActiveDocument.TrackRevisions = True
    ActiveDocument.TrackFormatting = True
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCrLf & .Description, vbCritical, "InsertPricingTableAndPaymentSchedule"
    End With
    Resume Exit_Here
End Sub

Public Sub InsertProducts()
    Dim fld As Field
    Dim rngStart As Range
    Dim rngEnd As Range
    Dim rngProducts As Range
    Dim intFldIndex As Integer
    Dim lngRangeStart As Long
    Dim lngRangeEnd As Long
    On Error GoTo ErrorHandler
    
    intFldIndex = 100
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldFormCheckBox Then
            fld.Select
            If fld.Index < intFldIndex Then
                lngRangeStart = Selection.Range.Start
                Set rngProducts = Selection.Range
            End If
            intFldIndex = fld.Index
            lngRangeEnd = Selection.Paragraphs(1).Range.End
        End If
    Next fld
    Selection.SetRange Start:=lngRangeStart, End:=lngRangeEnd
    Selection.Delete
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCrLf & .Description, vbCritical, "InsertProducts"
    End With
    Resume Exit_Here
End Sub

Public Sub PopulateWithXML()
    Dim dlg As FileDialog
    Dim varSelectedXMLFile As Variant
    Dim objContactXML As clsWordXMLData
    On Error GoTo ErrorHandler

    Set objContactXML = New clsWordXMLData
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    dlg.AllowMultiSelect = True
    dlg.InitialFileName = objContactXML.DefaultPath & "*.xml"
    If dlg.Show = -1 Then
        objContactXML.XMLFiles = varSelectedXMLFile
'        For Each varSelectedXMLFile In dlg.SelectedItems
'            objContactXML.XMLFile = varSelectedXMLFile
'        Next varSelectedXMLFile
        objContactXML.BindContentControlsv2
    End If

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCrLf & .Description, vbCritical, "PopulateWithXML"
    End With
    Resume Exit_Here
End Sub
