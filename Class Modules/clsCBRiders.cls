VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCBRiders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub CleanUpRiders()
'****************************************************************************************************
'Date Created:  02/09/2016
'Author:        Adrian A. Jones
'Class Name:    clsCBRiders
'Method:        CleanUpRiders
'
'Purpose:       1.  Deletes all the highlighted paragraphs with the names of riders and the fields,
'                   invisible to the naked eye, at the end of these paragraphs
'
'               2.  Also formats the rider's title so it automatically falls on the next page.
'
'               3.  Deletes the highlighted paragraph with the name of the active rider.
'                   Same as #1.
'
'               4.  And UNLINKS the field of the active rider.
'                   THIS is IMPORTANT because any changes made in the rider will be lost upon a
'                   field update, which can also happen when a document is printed out---if "Update
'                   Fields" on print is set. Of course pressing "F9" will totally wipe out any
'                   changes you've made if the rider is STILL a field.
'
'Revisions:     None as of 02/10/2016
'Notes:
'****************************************************************************************************

    Dim fld As Field
    Dim rngFld As Range
    Dim para As Paragraph
    Dim rngRiderHeader As Range
    Dim intRidersTotal As Integer
    
    Const strRIDER_NAME_TOKEN As String = "-"
    Const strRIDER_HEADER As String = "Schedule to College Board"
    
    On Error GoTo ErrorHandler
    
    If Not ActiveDocument.TrackRevisions Then
        ActiveDocument.TrackRevisions = True
    End If
    
    intRidersTotal = 0
    If ActiveDocument.Fields.Count <> 0 Then
        For Each fld In ActiveDocument.Fields
            'Each Rider is a separate paragraph with a field at the end
            'The purpose of the code is to determine if that field is expanded
            'to display the ENTIRE RIDER or if it is just a field at the end of
            'a lone paragraph.  If lone paragraph, then delete.
            Set para = fld.Result.Paragraphs(1)
            '"Riders" are IF fields in our CB contract template
            If (fld.Type = wdFieldIf) Then
                'Have to look at paragraph range in order to get to
                'the characters selection. I could have stuck .range at the end of Paragraphs(1)
                'above, however that line of code is already 3 dots deep (a reference of
                'a reference of a reference) and that's bad coding practice
                Set rngFld = para.Range
                rngFld.Select
                'Each rider name starts with a dash (-) and is highlighted
                'Added 02/11/2016, after testing against old riders, the OR clause.
                'May delete in the coming months when no old riders
                If ((rngFld.Characters(1).Text = strRIDER_NAME_TOKEN) _
                    And (rngFld.Characters(1).HighlightColorIndex = wdBrightGreen) _
                    And (Selection.Paragraphs.Count = 1)) _
                    Or ((rngFld.Characters(1).HighlightColorIndex = wdBrightGreen) _
                    And (Selection.Paragraphs.Count = 1)) Then
                    'Debug.Print "Selected Rider " & Selection.Text
                    Selection.Paragraphs(1).Range.Delete
                    intRidersTotal = intRidersTotal + 1
                Else
                    'Get rid of the highlighted Rider Names here
                    'During testing on 02/11/2016 noticed this ALONE works to clean up riders
                    Selection.Fields(1).Unlink
                    Selection.Paragraphs(1).Range.Delete
                    Selection.Find.Execute FindText:=strRIDER_HEADER
                    If Selection.Find.Found = True Then
                        Selection.ParagraphFormat.PageBreakBefore = True
                    End If
                End If
            End If
        Next fld
    End If
    
    '04/05/2017 Dale asked for us to remove MyRoad from the Rider sooo....
    '*******************************************************************************************
    If FindTextToDelete("MyRoad*and") Then
        MsgBox "Found MyRoad and Deleted", vbInformation, "Delete MyRoad Text"
    End If
    '*******************************************************************************************
    
    '04/06/2017 Removing the table surrounding Riders in K12 Documents
    '*******************************************************************************************
    FormatRiders
    '*******************************************************************************************
    
    '09/20/2017
    '*******************************************************************************************
    FormatCommonwealth_v2
    '*******************************************************************************************
    
    '08/1/2018 Should just CHANGE the link in the template, but today that takes too long so
    'I have to build a macro
    '*******************************************************************************************
    ReplaceIncorrectCBLink
    '*******************************************************************************************
    
    
    If intRidersTotal = 0 Then
        MsgBox "No Riders exist in this document: " & vbCrLf _
            & vbCrLf & "Currently active Document: " & ActiveDocument.Name & vbCrLf _
            & vbCrLf & "Make sure you are working in the correct document!", vbExclamation, "College Board Riders"
            If ActiveDocument.TrackRevisions Then
                ActiveDocument.TrackRevisions = False
            End If
    Else
        MsgBox "Number of Unnecessary Riders Found: " & intRidersTotal
    End If
    
Exit_Here:
    Set rngFld = Nothing
    Set para = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        .Source = "CleanUpRiders"
        .Raise .Number, .Source, .Description
    End With
End Sub

Private Function RiderIsActive(ByRef rrngCurrField As Range) As Boolean
    Dim intRidersTotal As Integer
    Const strRIDER_HEADER As String = "Schedule to College Board"
    Const strRIDER_NAME_TOKEN As String = "-"
    On Error GoTo ErrorHandler
    'Each rider name starts with a dash (-) and is highlighted
    'Added 02/11/2016, after testing against old riders, the OR clause.
    'May delete in the coming months when no old riders
    If ((rrngCurrField.Characters(1).Text = strRIDER_NAME_TOKEN) _
        And (rrngCurrField.Characters(1).HighlightColorIndex = wdBrightGreen) _
        And (Selection.Paragraphs.Count = 1)) _
        Or ((rrngCurrField.Characters(1).HighlightColorIndex = wdBrightGreen) _
        And (Selection.Paragraphs.Count = 1)) Then
        'Debug.Print "Selected Rider " & Selection.Text
        Selection.Paragraphs(1).Range.Delete
        intRidersTotal = intRidersTotal + 1
    Else
        'Get rid of the highlighted Rider Names here
        'During testing on 02/11/2016 noticed this ALONE works to clean up riders
        Selection.Fields(1).Unlink
        Selection.Paragraphs(1).Range.Delete
        Selection.Find.Execute FindText:=strRIDER_HEADER
        If Selection.Find.Found = True Then
            Selection.ParagraphFormat.PageBreakBefore = True
        End If
    End If
    
Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        .Source = "CleanUpRiders"
        .Raise .Number, .Source, .Description
    End With
End Function

Private Function FindTextToDelete(ByRef strTextToFind As String) As Boolean
    Dim blnFound As Boolean
    On Error GoTo ErrorHandler
    
    Selection.Find.ClearFormatting
    blnFound = Selection.Find.Execute(FindText:=strTextToFind, Forward:=True, MatchCase:=False, MatchWholeWord:=False, MatchWildcards:=True)
    If blnFound Then
        Selection.Delete
        FindTextToDelete = blnFound
    End If

Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FindTextToDelete"
    End With
    Resume Exit_Here
End Function

Private Sub FormatRiders()
    Dim blnFoundRider As Boolean
    Const strRIDER_TAG As String = "Schedule"
    
    On Error GoTo ErrorHandler
    
    'This must have Track Changes OFF in order to work. I think this may be the case for ANY search.
    If ActiveDocument.TrackRevisions Then
        ActiveDocument.TrackRevisions = False
    End If
    
    Selection.HomeKey Unit:=wdStory
    Do
        Selection.Find.ClearFormatting
        Selection.Find.Font.Bold = True
        With Selection.Find
            .Text = strRIDER_TAG
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        blnFoundRider = Selection.Find.Execute
        If blnFoundRider And Selection.Information(wdWithInTable) Then
            'remove table surrounding the rider
            'MsgBox "This is a Rider", vbInformation, "FormatRiders"
            RemoveSurroundingTables
            Selection.Paragraphs(1).Range.ParagraphFormat.PageBreakBefore = True
        End If
    Loop Until Not blnFoundRider

Exit_Here:
    If Not ActiveDocument.TrackRevisions Then
        ActiveDocument.TrackRevisions = True
    End If
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatRiders"
    End With
    Resume Exit_Here
End Sub

Private Sub ReplaceIncorrectCBLink()
    '8/1/2018
    Const strIncorrectLink As String = "http://research.collegeboard.org/data"
    Const strAUTOTEXT_NAME_CORRECTLINK As String = "HED Links-Use & Release of Data"
    
    Dim blnFound As Boolean
    
    On Error GoTo ErrorHandler
    
    Selection.HomeKey Unit:=wdStory
    Do
        blnFound = FindTextToDelete(strIncorrectLink)
        If blnFound Then
            Selection.TypeText Text:=Space(1)
            Application.Templates(ThisDocument.FullName).BuildingBlockEntries(strAUTOTEXT_NAME_CORRECTLINK).Insert _
                Where:=Selection.Range, RichText:=True
            Selection.Range.MoveEnd Unit:=wdParagraph, Count:=1
        End If
    Loop Until Not blnFound

Exit_Here:
    Exit Sub

ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "ReplaceIncorrectCBLink"
    End With
    Resume Exit_Here
End Sub
