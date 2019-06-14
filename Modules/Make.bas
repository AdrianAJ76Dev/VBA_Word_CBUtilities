Attribute VB_Name = "Make"
Option Explicit
Dim mobjContract As clsContract

Private Type CodeKeyBindings
    RoutineName As String
    FirstKey As Long
    SecondKey As Long
End Type

Public Sub Clean_Up_Contract()
    
    Const strREGEX_DATE_FORMAT_EBS As String = "" 'EBS = Enrollment Budget Schedule
    
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = wdAlertsNone
    Application.ScreenUpdating = False 'I may have to comment this out because I'm using the Selection Object and I think that NEEDS for Screenupdating to be TRUE
    Set mobjContract = New clsContract
    With ActiveDocument
        .TrackRevisions = True
        'Turn off seeing revisions or "Final" so the Find in the Clean_Up code works correctly
        .ShowRevisions = False
    End With
    
    With mobjContract
        .Clean_Up_Riders
        .FormatDates
        .Clean_Up_EnrollmentBudgetTable
        .CleanUp_General
    End With
    
    'Turn on seeing revisions or "Final: Show Markup"
    ActiveDocument.ShowRevisions = True
    
ExitHere:
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
    Set mobjContract = Nothing
    Exit Sub
    
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCrLf & .Source & vbCrLf & .Description, vbCritical + vbOKOnly, "Clean_Up_Contract"
    End With
    Resume ExitHere
End Sub

Public Sub Clean_Up_Riders()
    '************************************************************************************************
    'Author: Adrian Jones
    'Date Created: 08/24/2015
    'This routine exists in the class as well as here. The purpose for it HERE is to tie a keystroke
    'to this routine.
    'Alt C, R for C = Clean Up, R = Rider
    '************************************************************************************************
    
    Dim CBRiders As clsCBRiders
    On Error GoTo ErrorHandler
    
    Set CBRiders = New clsCBRiders
    CBRiders.CleanUpRiders
    
Exit_Here:
    Set CBRiders = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCrLf & .Description, vbCritical, "Clean_Up_Riders"
    End With
    Resume Exit_Here
End Sub

Public Sub Clean_Up_Riders_OLD()
'************************************************************************************************
'Author: Adrian Jones
'Date Created: 07/22/2015
'This routine exists in the class as well as here. The purpose for it HERE is to tie a keystroke
'to this routine.
'Alt C, R for C = Clean Up, R = Rider
'************************************************************************************************
    Dim fld As Field
    Dim rngRiderTitleParagraph As Range
    Dim rngRiderStart As Range
    Dim strRiderTitle As String
    Dim intRiderFieldCount As Integer
    Dim para As Paragraph
    Dim paras As Paragraphs
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    For Each fld In ActiveDocument.Fields
        'Only look at "Rider" Fields
        If fld.Type = wdFieldIf Then
            intRiderFieldCount = intRiderFieldCount + 1
            Set rngRiderTitleParagraph = fld.Result.Paragraphs(1).Range
            Set rngRiderStart = rngRiderTitleParagraph.Next(Unit:=wdParagraph, Count:=1)
            strRiderTitle = rngRiderTitleParagraph.Text
            If fld.Result.Paragraphs.Count = 1 Then
                'Debug.Print "Rider Title: " & strRiderTitle & "Paragraphs Count: " & fld.Result.Paragraphs.Count
            Else
                'Full justify all paragraphs
                Set paras = fld.Result.Paragraphs
                fld.Unlink
                For Each para In paras
                    para.Range.Select
                    If para.Alignment = wdAlignParagraphLeft Then
                        para.Alignment = wdAlignParagraphJustify
                    End If
                Next
                'Debug.Print "Active Rider: " & strRiderTitle & "With a Paragraph Count of: " & fld.Result.Paragraphs.Count
                'Debug.Print "UNLINKED FIELD/RIDER " & strRiderTitle
                'Place a page break before the rider
                rngRiderStart.ParagraphFormat.PageBreakBefore = True
            End If
            'Debug.Print
            rngRiderTitleParagraph.Delete
            'Debug.Print "DELETED FIELD/RIDER " & strRiderTitle
        End If
    Next fld
    'Debug.Print "Count of 'Rider Fields' " & intRiderFieldCount
    
ExitHere:
    Application.ScreenUpdating = True
    Set rngRiderTitleParagraph = Nothing
    Set rngRiderStart = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCrLf & .Description, vbCritical + vbOKOnly, "Clean_Up_Riders_Old"
    End With
    Resume ExitHere
End Sub

Public Sub Clean_Up_Riders_TheOnesWithTables()
    Dim objRiderTable As Table
    Dim rngOriginalSelection As Range
    Static intRoutineCount As Integer
    
    On Error GoTo ErrorHandler
    
    If intRoutineCount = 0 Then
        'Set Revision State to Final
        With ActiveDocument
            .TrackRevisions = True
            'Turn off seeing revisions or "Final" so the Find in the Clean_Up code works correctly
            .ShowRevisions = False
        End With
        intRoutineCount = intRoutineCount + 1
        Set rngOriginalSelection = Selection.Range
    End If
    
    'See if you are in a table 1st
    If Selection.Information(wdWithInTable) Then
        Set objRiderTable = Selection.Tables(1)
        objRiderTable.Select
        objRiderTable.ConvertToText Separator:=wdSeparateByParagraphs, NestedTables:=False
        rngOriginalSelection.Select
        Set objRiderTable = Nothing
        'Call it again because we are a table within a table
        Clean_Up_Riders_TheOnesWithTables
    Else
        MsgBox "You're not within a table. Can't run code", vbExclamation, "Clean Up_Riders_TheOnesWithTables"
    End If
    
    intRoutineCount = 0
    'Turn on seeing revisions or "Final: Show Markup"
    ActiveDocument.ShowRevisions = True
    
ExitHere:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCrLf & .Description, vbCritical + vbOKOnly, "Clean_Up_Riders_TheOnesWithTables"
    End With
    Resume ExitHere
End Sub

Public Sub MakeAmendment()
    Const FULL_NAME_TEMPLATE As String = "K12 Amendment Template v3 Final.dotx"
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Set mobjContract = New clsContract
    
    'REMEMBER TO REDO THIS BY TAKING THE FORM OUT OF THIS CLASS (Assbackwards)
    'Determine if K12 or HED contract that is currently active
    frmAmendment.Show
        
    'If HED run CreateAmendmentFromMain like I have been
    If frmAmendment.Tag = "K12" Then
        mobjContract.CreateK12AmendmentfromTemplate PullFromSharePoint(FULL_NAME_TEMPLATE)
    ElseIf frmAmendment.Tag = "HED" Then
        mobjContract.CreateAmendmentFromMain
    End If
    
    Unload frmAmendment


ExitHere:
    Application.ScreenUpdating = True
    Set mobjContract = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCrLf & .Description, vbCritical + vbOKOnly, "MakeAmendment"
    End With
    Resume ExitHere
End Sub

Public Sub RemoveSurroundingTables()
'**********************************************************************************************************************************************************************************
' Keyboard Shortcut Alt-R,T
' Important Note: For some reason, I have to be careful what module I put my code in
'**********************************************************************************************************************************************************************************
    On Error GoTo ErrorHandler
    
    Do While Selection.Information(wdWithInTable) = True
        Selection.Rows.ConvertToText Separator:=wdSeparateByParagraphs, NestedTables:=False
    Loop
    
    'Get's rid of the additional paragraph spacing because the conversion gives you extra spacing (probably a style would fix this)
    Selection.ParagraphFormat.SpaceAfter = 0

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "RemoveSurroundingTables"
    End With
    Resume Exit_Here
End Sub

Public Sub FormatRiders()
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

Public Sub MoveToOuterMostTable()
    Dim lngStartNestLevel As Long
    On Error GoTo ErrorHandler
    
    If Selection.Information(wdWithInTable) Then
        If Selection.Tables(1).NestingLevel > 1 Then
            Do
                Selection.MoveUp Unit:=wdParagraph
            Loop Until Selection.Tables(1).NestingLevel = 1
        End If
    End If

Exit_Here:
    Exit Sub

ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "MoveToOuterMostTable"
    End With
    Resume Exit_Here
End Sub


Private Sub FindScheduleWithParagraph()
'
' FindScheduleWithParagraph Macro
'
'
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "Schedule^p"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
End Sub

Public Sub RefreshShortcuts()
    Dim i As Integer
    Dim aKey As KeyBinding
    Dim intShortCutKeyCount As Integer
    Dim lngKeyCode As Long
    Dim lngMsgBoxResult As Long
    Dim udtCodeShortCuts() As CodeKeyBindings
    
    On Error GoTo ErrorHandler

    'Give the user the opportunity to cancel this procedure
    
    lngMsgBoxResult = MsgBox("Warning!! This will REPLACE any Shortcut Keys you have assigned" & vbCr _
                        & "with Shortcut Keys from the CMUtilities Template/Add-In." & vbCr _
                        & "So any personal macros you have created with shortcut keys may no longer work using those keystorkes." & vbCr _
                        & "Do you still wish to refresh the Shortcut Keys for the CMUtility Macros?", vbExclamation + vbYesNoCancel, "Refresh Shortcut Keys")
                        
    If lngMsgBoxResult = vbYes Then
        'Retrieve ShortCut Keys from Template
        Debug.Print ThisDocument.Name
        'Application.CustomizationContext = ThisDocument
        CustomizationContext = ThisDocument
        'This is probably a one time only thing, but I had double the keybindings with HALF having NO Command attached to it
        'I should be able to remove this code from here as I don't believe I have to check this all the time
        'This code is possibly JUST Utility code.
        For Each aKey In KeyBindings
            If Len(aKey.Command) = 0 Then
                aKey.Clear
            End If
        Next aKey
        
        i = 0
        intShortCutKeyCount = KeyBindings.Count - 1
    
    '    Debug.Print "Count of KeyBindings = " & KeyBindings.Count
        If intShortCutKeyCount <> 0 Then
            ReDim udtCodeShortCuts(intShortCutKeyCount)
            'Pick up and store all ShortCut keys from the template
            For Each aKey In KeyBindings
                udtCodeShortCuts(i).RoutineName = aKey.Command
                udtCodeShortCuts(i).FirstKey = aKey.KeyCode
                udtCodeShortCuts(i).SecondKey = aKey.KeyCode2
                i = i + 1
'                Debug.Print "Adding..." & aKey.Command
'                Debug.Print aKey.KeyString
'                Debug.Print "KeyCode=" & aKey.KeyCode
'                Debug.Print "KeyCode2=" & aKey.KeyCode2
'                Debug.Print
            Next aKey
    
    '        Debug.Print
            CustomizationContext = NormalTemplate
            KeyBindings.ClearAll
            For i = 0 To UBound(udtCodeShortCuts())
    '            Debug.Print "Adding..." & udtCodeShortCuts(i).RoutineName
    '            KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    '                Command:=udtCodeShortCuts(i).RoutineName, _
    '                KeyCode:=BuildKeyCode(wdKeyAlt, udtCodeShortCuts(i).FirstKey), _
    '                KeyCode2:=BuildKeyCode(udtCodeShortCuts(i).SecondKey)
                
                KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                    Command:=udtCodeShortCuts(i).RoutineName, _
                    KeyCode:=udtCodeShortCuts(i).FirstKey, _
                    KeyCode2:=udtCodeShortCuts(i).SecondKey
            Next i
        End If
        NormalTemplate.Save
        MsgBox "Shortcut Keys have been refreshed", vbInformation, "RefreshShortcuts"
    Else
        MsgBox "Refresh Shortcut Keys operation canceled.", vbInformation, "Refresh Shortcut Keys"
    End If
    
Exit_Here:
    Exit Sub

ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "RefreshShortcuts"
    End With
    Resume Exit_Here
End Sub

Public Sub CreateCoversheetForSignatureForSpecificContract()
    Dim dsMain As MailMergeDataSource
    Dim numRecord As Integer
    Dim ContractNumber As String
    Const strMergeField As String = "Contract Number"
    
    On Error GoTo ErrorHandler
    
    ActiveDocument.MailMerge.ViewMailMergeFieldCodes = False
    Application.ScreenRefresh
    Application.ScreenUpdating = True
    ContractNumber = InputBox("Enter or copy and past contract number", "Contract Number")
    
    Set dsMain = ActiveDocument.MailMerge.DataSource
    dsMain.ActiveRecord = wdFirstRecord
    If dsMain.FindRecord(FindText:=ContractNumber, Field:="Contract_Number") = True Then
        numRecord = dsMain.ActiveRecord
    Else
        MsgBox "Something went wrong"
    End If

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Create Coversheet For Signature For Specific Contract"
    End With
    Resume Exit_Here
End Sub

Public Function TwoWeeksFromToday()
    On Error GoTo ErrorHandler
    
    TwoWeeksFromToday = Format(Now() + 14, "mm/dd/yyyy")

Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Two Weeks From Today"
    End With
    Resume Exit_Here
End Function

Public Function PullFromSharePoint(ByVal vstrFileName As String) As String
    '01/31/2018 Retrieve the template from SharePoint
    Const mstrPATH_SHAREPOINT As String = "https://mysharepoint/sdp/oosvp/RAS/Contracts Management/"
    
    '05-31-2019 Retrieve from the template directory
    Const mstrPATH_SHAREPOINT_WRD_TEMPLATES As String = "https://mysharepoint/sdp/oosvp/RAS/Contracts Management/CM-transfer/Templates/"
    
    PullFromSharePoint = mstrPATH_SHAREPOINT_WRD_TEMPLATES & vstrFileName
    
Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "SoleSourceTemplateFullName"
    End With
    Resume Exit_Here
End Function

