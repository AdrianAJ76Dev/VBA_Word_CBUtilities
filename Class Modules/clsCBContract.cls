VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCBContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*********************************************************************************************
'Date Created:  08/05/2015
'Author:        Adrian A. Jones
'Class Name:    clsCBContract
'
'Purpose:       1. To model College Board standard contracts to improve processing them and
'               to better understand what this model will look like as a Salesforce object.
'
'               2. It's purpose also is to 'model the document clean-up' we have to do because
'               of the less than robust contract generation we face today
'
'Revisions:     None as of 08/05/2015
'Notes:         I'm already thinking the Contracts class should be an abstract class.
'               So already I've changed the name to 'clsCBContract' from 'clsContract_v2'
'               Bad naming anyway hardcoding the version number in the name. How would I
'               update the class and remember truly what version it is if I include the
'               version number in the name of the class? "Oh it's version 6 really even
'               if the class name is clsContract_v2!"
'
'               'CB' stands for 'College Board'
'               IMPORTANT: NOTHING is being inserted into the contract!  It is being cleaned
'               up and properly formatted.  It is NOT being constructed!!
'
'               Previous code, clsContract, used the Selection Object to make edits and
'               changes. This version I want to use ranges.
'*********************************************************************************************

'The Enums
'QUESTION:  Do I store these as DocVars, as DocProps or do I 'READ' the contract and determine
'           that?
Private Enum ContractType
    HED = 1
    K12 = 2
End Enum

Private Enum ContractSection
    Standard = 1
    Amendment = 2
    SignaturePage = 3
    ClientInformationPage = 4
    Rider = 5
    BudgetSchedule = 6
End Enum

Private Enum ContractState
    NewlyGeneratedFromSalesforce = 1
    Redlined = 2
    RedlinedFinal = 3
    FinalReadyToSend = 4
End Enum

Private Enum BudgeScheduleCols
    ProductName = 1
    StartDate = 2
    EndDate = 3
    Quantity = 4
    UnitPrice = 5
    Cost = 6
    Discount = 7
    TotalCost = 8
End Enum

'SF generated contracts are 'Word IF fields' where the appropriate rider is displayed within that IF field
Private mobjRider As Field
Private mparaHighlightedRiderTitle As Paragraph
Private mintRiderFieldParagraphCount As Integer
Private mdatTermDateStart As Date
Private mdatTermDateEnd As Date

'The idea here is to search within the scope of the rider ONLY to determine what is a start date and an end date
Private mrngRiderRange As Range
Private mintRiderCount As Integer
Private mtblBudgetSchedule As Table

'Not implementing right away, but for later contract verification after generation
'May want to check if term dates in the riders and in the budget schedule match
'May want to check if the riders were inserted and throw up a message box when they weren't
Private mstrContractGenerationErrors() As String
Private mlngIfField As Long 'For storing wdFieldIf value which is a global enumeration value
Private Const mblnDONT_SHOW_REVISIONS As Boolean = False

'*********************************************************************************************
'Date Created:  09/18/2015
'Author:        Adrian A. Jones
'Class Name:    clsCBContract
'Routine Name:  DetermineRiders
'
'Purpose:       1. Remove all unnecessary rider fields and rider headings
'
'Revisions:     None as of 09/18/2015
'
'Notes:         The reality of contracts generation today is the clean up of riders.
'               ALL riders are included in the agreement and "Turned On" or "Turned Off"
'               in each document where it is up to the contract specialist to DELETE the
'               unnecessary riders.
'*********************************************************************************************
Public Sub DetermineActiveRiders()
    Dim fld As Field
    Dim FieldCount As Integer
    
    On Error GoTo ErrorHandler
    
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.Fields.Update
    FieldCount = 0
    
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldIf And fld.Result.Paragraphs.Count > 1 Then
            'For debugging
            '**********************************************************************************************************
            FieldCount = FieldCount + 1
    
            Debug.Print "Field #: " & FieldCount & " 1st Paragraph Text/Title: " & fld.Result.Paragraphs(1).Range.Text
            
            fld.Code.Paragraphs(1).Range.Select
            Debug.Print "Code Paragraph Count is: " & fld.Code.Paragraphs.Count
            
            fld.Result.Paragraphs(1).Range.Select
            Debug.Print "Result Paragraph Count is: " & fld.Result.Paragraphs.Count
            Debug.Print
            '**********************************************************************************************************
        End If
    Next fld
    
    Debug.Print
    Debug.Print "Field Count for Riders = " & FieldCount

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "DetermineActiveRiders"
    End With
    Resume Exit_Here
End Sub

Public Sub ActivateRider()
    Dim rngRider As Range
    Dim rngSelectPastedRider As Range
    Dim intNestingLevel As Integer
    
    Const strIF_FIELD_TEXT As String = "False"
    
    On Error GoTo ErrorHandler
    
    
    Set rngRider = Selection.Fields(1).Code
    Set rngSelectPastedRider = Selection.Range
    rngRider.Copy
    rngSelectPastedRider.Paste
    
    If rngSelectPastedRider.Fields.Count > 0 Then
        rngSelectPastedRider.Fields.Update
    End If
    
    If InStr(1, Selection.Paragraphs(1).Range.Text, strIF_FIELD_TEXT) <> 0 Then
        Selection.Paragraphs(1).Range.Delete
    End If
    
    'If rider is imprisioned, release rider from table(s)
'    rngSelectPastedRider.Select
'    If Selection.Tables.Count > 0 Then
'        Dim intNestingLevel As Integer
'        Selection.GoTo What:=wdTable, Which:=wdGoToNext
'        Selection.Collapse Direction:=wdCollapseStart
'        Do
'            Selection.MoveDown Unit:=wdParagraph
'            If Selection.Information(wdWithInTable) Then
'
'            End If
'            intNestingLevel
'    End If
    
    Selection.MoveDown Unit:=wdParagraph, Count:=5
    Do
        If Selection.Information(wdWithInTable) Then
            Selection.Tables(1).ConvertToText Separator:=wdSeparateByParagraphs, NestedTables:=False
            Selection.MoveDown Unit:=wdParagraph, Count:=1
        End If
    Loop Until Selection.Information(wdWithInTable) = False
    rngSelectPastedRider.ParagraphFormat.SpaceAfter = 0
    
'    Do
'        If Selection.Tables.Count > 0 Then
'            Selection.Tables(1).Range.Select
'            Selection.Collapse wdCollapseStart
'        End If
'    Loop Until Selection.Information(wdWithInTable) = False
'
'    Do
'        If Selection.Information(wdWithInTable) Then
'            Selection.Tables(1).ConvertToText Separator:=wdSeparateByParagraphs, NestedTables:=False
'        End If
'    Loop Until Selection.Information(wdWithInTable) = False

Exit_Here:
    Set rngRider = Nothing
    Set rngSelectPastedRider = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "ActivateRider"
    End With
    Resume Exit_Here
End Sub

Public Sub DeleteUnncessaryRiders()
    Dim fld As Field
    Dim rng As Range
    On Error GoTo ErrorHandler
    
    ActiveDocument.Fields.Update
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldIf Then
            Set rng = fld.Code.Paragraphs(1).Range
            rng.Move Unit:=wdWord, Count:=-1
            If Selection.Range.HighlightColorIndex = wdNoHighlight Then
                'We want to keep the fields that are not highlighted
            Else
                fld.Delete
                rng.Paragraphs(1).Range.Select
                Selection.Delete
            End If
        End If
    Next fld
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "ActivateRider"
    End With
    Resume Exit_Here
End Sub

