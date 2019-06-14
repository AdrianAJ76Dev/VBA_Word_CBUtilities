Attribute VB_Name = "modCommonTools"
Option Explicit

Public Sub ReverseName()
    On Error GoTo ErrorHandler
    
    Dim varManagerName As Variant
    varManagerName = Split(Selection, " ")
    Selection = varManagerName(1) + Space(1) + varManagerName(0)
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "ReverseName"
    End With
    Resume Exit_Here
End Sub

Public Sub TestFileSaveAsDialog()
    Dim strPATH As String
    On Error GoTo ErrorHandler

    strPATH = "Q:\RAS Contracts Management\Pivotal\K12 Contracts - Pivotal\_PSAT Agreements\OH\Cincinnati\Agreements\2013-2014\Draft"
    With Application.FileDialog(msoFileDialogSaveAs)
        .InitialFileName = strPATH
        .Show
    End With

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "TestFileSaveAsDialog"
    End With
    Resume Exit_Here
End Sub

Public Function SpellNumber(ByVal MyNumber) As String
    Dim Temp As String
    Dim DecimalDigits As String
    Dim NumberAsString As String
    Dim DecimalPlace, Count
    
    On Error GoTo ErrorHandler
    
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "

    'String representation of amount.
    If IsNumeric(MyNumber) Then
        MyNumber = Trim(Str(MyNumber))

        'Position of decimal place 0 if none.
        DecimalPlace = InStr(MyNumber, ".")
    
        'Convert cents and set MyNumber to dollar amount.
        If DecimalPlace > 0 Then
            DecimalDigits = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
        End If
        
        Count = 1
        Do While MyNumber <> ""
            Temp = GetHundreds(Right(MyNumber, 3))
            If Temp <> "" Then NumberAsString = Temp & Place(Count) & NumberAsString
            If Len(MyNumber) > 3 Then
                MyNumber = Left(MyNumber, Len(MyNumber) - 3)
            Else
                MyNumber = ""
            End If
            Count = Count + 1
        Loop
    Else
        MsgBox "The current selection is not a number or part of it contains letters." & vbCr _
        & "Please select only numbers to be converted.", vbExclamation, "Spell Number"
        NumberAsString = Trim(MyNumber)
    End If
        
    SpellNumber = NumberAsString

Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "SpellNumber"
    End With
    Resume Exit_Here
End Function

'*******************************************

' Converts a number from 100-999 into text *

'*******************************************

Function GetHundreds(ByVal MyNumber)
    Dim Result As String

    On Error GoTo ErrorHandler

    If Val(MyNumber) = 0 Then Exit Function

    MyNumber = Right("000" & MyNumber, 3)



    ' Convert the hundreds place.

    If Mid(MyNumber, 1, 1) <> "0" Then

        Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "

    End If



    ' Convert the tens and ones place.

    If Mid(MyNumber, 2, 1) <> "0" Then

        Result = Result & GetTens(Mid(MyNumber, 2))

    Else

        Result = Result & GetDigit(Mid(MyNumber, 3))

    End If



    GetHundreds = Result

Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "GetHundreds"
    End With
    Resume Exit_Here
End Function

'*********************************************

' Converts a number from 10 to 99 into text. *

'*********************************************

Function GetTens(TensText)

    Dim Result As String

    On Error GoTo ErrorHandler

    Result = ""           ' Null out the temporary function value.

    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...

        Select Case Val(TensText)

            Case 10: Result = "Ten"

            Case 11: Result = "Eleven"

            Case 12: Result = "Twelve"

            Case 13: Result = "Thirteen"

            Case 14: Result = "Fourteen"

            Case 15: Result = "Fifteen"

            Case 16: Result = "Sixteen"

            Case 17: Result = "Seventeen"

            Case 18: Result = "Eighteen"

            Case 19: Result = "Nineteen"

            Case Else

        End Select

    Else                                 ' If value between 20-99...

        Select Case Val(Left(TensText, 1))

            Case 2: Result = "Twenty "

            Case 3: Result = "Thirty "

            Case 4: Result = "Forty "

            Case 5: Result = "Fifty "

            Case 6: Result = "Sixty "

            Case 7: Result = "Seventy "

            Case 8: Result = "Eighty "

            Case 9: Result = "Ninety "

            Case Else

        End Select

        Result = Result & GetDigit(Right(TensText, 1))  ' Retrieve ones place.

    End If

    GetTens = Result

Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "GetTens"
    End With
    Resume Exit_Here
End Function

'*******************************************

' Converts a number from 1 to 9 into text. *

'*******************************************

Function GetDigit(Digit)
    
    On Error GoTo ErrorHandler
    Select Case Val(Digit)

        Case 1: GetDigit = "One"

        Case 2: GetDigit = "Two"

        Case 3: GetDigit = "Three"

        Case 4: GetDigit = "Four"

        Case 5: GetDigit = "Five"

        Case 6: GetDigit = "Six"

        Case 7: GetDigit = "Seven"

        Case 8: GetDigit = "Eight"

        Case 9: GetDigit = "Nine"

        Case Else: GetDigit = ""

    End Select

Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "GetDigit"
    End With
    Resume Exit_Here
End Function

Function SpellNumberOutToCurrency(ByVal MyNumber)
    Dim Dollars, Cents, Temp
    Dim DecimalPlace, Count

    On Error GoTo ErrorHandler

    ReDim Place(9) As String

    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "



    ' String representation of amount.

    MyNumber = Trim(Str(MyNumber))



    ' Position of decimal place 0 if none.

    DecimalPlace = InStr(MyNumber, ".")

    ' Convert cents and set MyNumber to dollar amount.

    If DecimalPlace > 0 Then

        Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))

        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))

    End If



    Count = 1

    Do While MyNumber <> ""

        Temp = GetHundreds(Right(MyNumber, 3))

        If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars

        If Len(MyNumber) > 3 Then

            MyNumber = Left(MyNumber, Len(MyNumber) - 3)

        Else

            MyNumber = ""

        End If

        Count = Count + 1

    Loop



    Select Case Dollars

        Case ""

            Dollars = "No Dollars"

        Case "One"

            Dollars = "One Dollar"

        Case Else

            Dollars = Dollars & " Dollars"

    End Select



    Select Case Cents

        Case ""

            Cents = " and No Cents"

        Case "One"

            Cents = " and One Cent"

        Case Else

            Cents = " and " & Cents & " Cents"

    End Select



    SpellNumberOutToCurrency = Dollars & Cents

Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "SpellNumberOutToCurrency"
    End With
    Resume Exit_Here
End Function

Public Sub CopyDocumentProperties()
    Dim tmplCopyFrom As Template
    Dim docCopyFrom As Document
    Dim tmplCopyTo As Template
    Dim strTemplatePath As String
    Dim strTemplateCopyFrom As String
    Dim objCustDocProp As DocumentProperty
    
    Const strTEMPLATE_NAME_HED_RP_TERMS As String = "HED - RP Terminations.dotx"
    Const strPATH_TEMPLATE As String = "C:\Documents and Settings\ajones\Application Data\Microsoft\Templates\"
    
    On Error GoTo ErrorHandler
    
    'May need to throw this into the global template
    Set tmplCopyFrom = ThisDocument.AttachedTemplate
    
    If ActiveDocument.Name <> ThisDocument.Name Then
        For Each objCustDocProp In ThisDocument.CustomDocumentProperties
            If objCustDocProp.Name <> "ExecutionDate" Then
                ActiveDocument.CustomDocumentProperties.Add _
                    Name:=objCustDocProp.Name, LinkToContent:=False, Value:=objCustDocProp.Value, _
                    Type:=msoPropertyTypeString
            End If
        Next objCustDocProp
    End If
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "CopyDocumentProperties"
    End With
    Resume Exit_Here
End Sub

Public Sub FormatDateSpellOutMonth() 'Shortcut Key: Alt-F, D "D" for Date
    On Error GoTo ErrorHandler
    
    'Use a regular expression here
    If IsDate(Selection.Text) Then 'May want to use regular expressions to verify the pattern
        Selection.Text = Format(Selection, "mmmm d, yyyy")
    Else
        MsgBox "A Date doesn't appear to be selected", vbExclamation, "Format Date Spell Out Month"
    End If

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatDateSpellOutMonth"
    End With
    Resume Exit_Here
End Sub

Public Sub FormatPhoneNumber() 'Shortcut Key: Alt-F, P "P" for Phone
    Dim strPhoneNumberDigits As String
    On Error GoTo ErrorHandler
    
    'Check selection to see if 10 digits are selected
    strPhoneNumberDigits = Trim(Selection.Words(1).Text)
    If IsNumeric(strPhoneNumberDigits) And Len(strPhoneNumberDigits) = 10 Then
        strPhoneNumberDigits = Format(strPhoneNumberDigits, "(###)-###-####")
        Selection.Words(1).Text = strPhoneNumberDigits
    Else
        MsgBox "Your selection does not solely consist of numbers " & vbCr & _
        "or consists of more than or less than 10 digits " & "Number Count: " & Len(strPhoneNumberDigits), _
        vbExclamation, "Convert Number to Phone Format"
    End If

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatPhoneNumber"
    End With
    Resume Exit_Here
End Sub

Public Function FirstSecondThirdForth(ByRef strDigit As String) As String
    Dim strLastDigit As String
    On Error GoTo ErrorHandler
    
    strLastDigit = Right(Trim(strDigit), 1)
    Select Case strLastDigit
        Case 1
            strDigit = strDigit & "st"
            
        Case 2
            strDigit = strDigit & "nd"
            
        Case 3
            strDigit = strDigit & "rd"
            
        Case Else
            strDigit = strDigit & "th"
        
    End Select
    
    FirstSecondThirdForth = strDigit
    
Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FirstSecondThirdForth"
    End With
    Resume Exit_Here
End Function

Public Function FormatDateForAmendment(ByRef strAmendmentDate As String) As String
    Dim datAmendmentDate As Date
    On Error GoTo ErrorHandler
    
    'Grab the day for 1st, 2nd, 3rd, 4th formatting
    datAmendmentDate = strAmendmentDate
    FormatDateForAmendment = FirstSecondThirdForth(Day(strAmendmentDate)) & " day of " & Format(datAmendmentDate, "mmmm") & ", " & Format(datAmendmentDate, "yyyy")
    
Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatDateForAmendment"
    End With
    Resume Exit_Here
End Function

Public Function GetAmendmentDate() As String
    Dim strEnteredDate As String
    On Error GoTo ErrorHandler
    
    strEnteredDate = InputBox("Enter Date of the Main Agreement in m/dd/yy", "Main Agreement Date")
    GetAmendmentDate = FormatDateForAmendment(strEnteredDate)
    
Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "GetAmendmentDate"
    End With
    Resume Exit_Here
End Function

Public Sub RunInputAmendmentDate()
    On Error GoTo ErrorHandler
    
    Selection.Text = GetAmendmentDate

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "RunInputAmendmentDate"
    End With
    Resume Exit_Here
End Sub

'*******************************************

' Strips the leading space from a sentence *
' Alt-S-S or ASS Alt S: Strip S: Space     *

'*******************************************
Public Sub StripSpace()
    Dim rngStripSpaceSentence As Range
    On Error GoTo ErrorHandler
    
    Selection.Sentences(1).Select
    Set rngStripSpaceSentence = Selection.Range
    rngStripSpaceSentence.Text = Trim(rngStripSpaceSentence.Text)

Exit_Here:
    Set rngStripSpaceSentence = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "StripSpace"
    End With
    Resume Exit_Here
End Sub

Public Sub FormatPrice()
    '05/08/2017 Figure out what the keystroke/command that changes the selection from an entire cell to
    'the word selected?
    'Recorded the keystroke and it is:
    'Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    Dim Price As String
    
    On Error GoTo ErrorHandler
    If IsNumeric(Format(Selection.Text, "#")) Then
        If InStr(1, Selection.Text, "$") = 0 Then
            Price = Selection.Text
        Else
            Price = Val(Mid(Selection.Text, 2))
        
        End If
        Selection.Text = Format(Price, "$ ###,###,###.00")
    Else
        MsgBox "Selection does not appear to be a number", vbInformation, "FormatPrice"
    End If
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatPrice"
    End With
    Resume Exit_Here
End Sub

Public Sub FormatClientInfoTable()
    Dim tbl As Table
    On Error GoTo ErrorHandler
    
    Set tbl = Selection.Tables(1)
    
    With tbl
        .TopPadding = InchesToPoints(0)
        .BottomPadding = InchesToPoints(0)
        .LeftPadding = InchesToPoints(0.05)
        .RightPadding = InchesToPoints(0.05)
        .Spacing = 0
        .AllowPageBreaks = True
        .AllowAutoFit = False
        .Rows.LeftIndent = InchesToPoints(0.5)
    End With
    tbl.AutoFitBehavior wdAutoFitContent

Exit_Here:
    Set tbl = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatClientInfoTable"
    End With
    Resume Exit_Here
End Sub

Public Sub RemoveTermDatesFromFeeSchedule()
    Dim tbl As Table
    Dim tblNewFromSplit As Table
    Dim currRow As Row
    Dim c As Cell
    Dim rowHeaderRowColumnCount As Integer
    Dim rngSceneOfTheCrime As Range 'Paragraph/Range where the Split occurred.  By deleting it, Table is merged together like before.
    
'    Const COL_WIDTH_PRODUCT_NAME As Single = 5.16!
'    Const COL_WIDTH_QUANTITY As Single = 0.62!
'    Const COL_WIDTH_TOTAL_COST As Single = 0.87!

    Const BMK_NAME_PARAGRAPH_SPLIT As String = "SplitTableParagraph"
    Const BMK_NAME_SPLIT_TABLE2 As String = "SplitTable2"
    
    On Error GoTo ErrorHandler
    
    If Selection.Information(wdWithInTable) Then
        Application.DisplayAlerts = wdAlertsNone
        Application.ScreenUpdating = False
        
        Set tbl = Selection.Tables(1)
        tbl.AllowAutoFit = False
        
        rowHeaderRowColumnCount = tbl.Rows(1).Cells.Count
        
        'Take the 1st rows count of columns --- that's the header --- and cycle through the rows until reaching a row that
        'has a DIFFERENT column count
        For Each currRow In tbl.Rows
            If currRow.Cells.Count <> rowHeaderRowColumnCount Then
                currRow.Select
                
                'Get reference to bottom half of table because of split by adding a bookmark
                Selection.Bookmarks.Add Name:=BMK_NAME_SPLIT_TABLE2, Range:=Selection.Range
                Selection.SplitTable
                Selection.Bookmarks.Add Name:=BMK_NAME_PARAGRAPH_SPLIT, Range:=Selection.Range
                
                Set tblNewFromSplit = ActiveDocument.Bookmarks(BMK_NAME_SPLIT_TABLE2).Range.Tables(1)
                'Test that the bookmark works
'                tblNewFromSplit.Select
                
                'The original table selected, before the split, is the "top half" of the table
                tbl.Columns(2).Select
                tbl.Columns(2).Delete
                tbl.Columns(2).Select
                tbl.Columns(2).Delete
                
                For Each c In tbl.Columns(2).Cells
                    c.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                Next c
                
                ActiveDocument.Bookmarks(BMK_NAME_SPLIT_TABLE2).Delete
                
                'Old method of resizing table parts
                'This changes the size of the table at each assignment
'                tbl.Columns(1).Width = InchesToPoints(COL_WIDTH_PRODUCT_NAME)
'                tbl.Columns(2).Width = InchesToPoints(COL_WIDTH_QUANTITY)
'                tbl.Columns(3).Width = InchesToPoints(COL_WIDTH_TOTAL_COST)
                
                'New method of resizing table parts
                'Resize Table Top - Table 1 of split, to size of window
                tbl.AutoFitBehavior wdAutoFitWindow
                
                'Resize Table bottom - Table 2 of split, to size of window
                tblNewFromSplit.AutoFitBehavior wdAutoFitWindow
                
                
                tbl.AllowAutoFit = False
                tblNewFromSplit.AllowAutoFit = False
                ActiveDocument.Bookmarks(BMK_NAME_PARAGRAPH_SPLIT).Range.Delete
                ActiveDocument.Bookmarks(BMK_NAME_PARAGRAPH_SPLIT).Delete
                               
                'Reinforce resize of entire table
                tbl.AutoFitBehavior wdAutoFitWindow
                
                Exit For
            End If
        Next currRow
    Else
        MsgBox "You have to 1st put the cursor into a table", vbInformation, "Remove Term Date Columns in Fee Schedule"
    End If
    
Exit_Here:
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
    Set tbl = Nothing
    Set rngSceneOfTheCrime = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatClientInfoTable"
    End With
    Resume Exit_Here
End Sub

Public Sub FindTextToDelete(ByRef strTextToFind As String)
    Dim blnFound As Boolean
    Dim intDeleteCount As Integer
    
    Const strDELETED_TEXT As String = "MyRoad" 'Should not hardcode this but for now it's fine.  Should be whatever's passed in, but "MyRoad*and" may cause confusion.
    
    On Error GoTo ErrorHandler
    
    If Not ActiveDocument.TrackRevisions Then
        ActiveDocument.TrackRevisions = True
    End If
    intDeleteCount = 0
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Do
        blnFound = Selection.Find.Execute(FindText:=strTextToFind, Forward:=True, MatchCase:=False, MatchWholeWord:=False, MatchWildcards:=True)
        If blnFound Then
            Selection.Delete
            intDeleteCount = intDeleteCount + 1
        End If
    Loop Until Not blnFound
    
    If intDeleteCount <> 0 Then
        MsgBox "Deleted " & intDeleteCount & " occurrences of " & strDELETED_TEXT, vbInformation, "Delete " & strDELETED_TEXT
    End If

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FindTextToDelete"
    End With
    Resume Exit_Here
End Sub

Public Function FindBasedOnPattern(ByVal vstrPattern As String) As Variant
    Dim blnFoundText As Boolean
    Dim astrFoundText() As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Selection.HomeKey Unit:=wdStory
    i = 0
    Do
        blnFoundText = Selection.Find.Execute(FindText:=vstrPattern, Forward:=True, MatchWildcards:=True)
        If blnFoundText Then
            ReDim Preserve astrFoundText(i)
            astrFoundText(i) = Selection.Text
        End If
        Selection.Collapse Direction:=wdCollapseEnd
    Loop Until Not blnFoundText
    
    If UBound(astrFoundText) > 0 Then
        FindBasedOnPattern = astrFoundText()
    End If
    
Exit_Here:
    Exit Function
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FindBasedOnPattern"
    End With
    Resume Exit_Here
End Function

Public Sub FormatCommonwealth()
    '03/13/2018 I may retire these constants in favor of a search for "State of [STATE]
    Const PARA_COMMONWEALTH As Integer = 3
    
    '03/13/2018
    Const TARGET_TEXT_STATE As String = "State"
    Const REPLACE_TEXT_COMMONWEALTH As String = "Commonwealth"
    
    '03/12/2018
    Dim ingWordLast As Integer
    Dim wrdLast As Range
    Dim rngReplaceStateOfState As Range
    
    Dim rngFirstParaOfContract As Range
    Dim wrdState As Range
    Dim rngSelectedStateOfPhrase As Range
    Dim strState As String
    Dim IsCommonWealth As Boolean
    Dim aCommonwealths() As Variant
    Dim aResult() As String
    
    On Error GoTo ErrorHandler
    
    Selection.HomeKey Unit:=wdStory
    IsCommonWealth = False
    aCommonwealths = Array("Kentucky", "Massachusetts", "Pennsylvania", "Virginia")
    
    Set rngFirstParaOfContract = ActiveDocument.Paragraphs(PARA_COMMONWEALTH).Range
    ingWordLast = rngFirstParaOfContract.Words.Count
    Set wrdLast = rngFirstParaOfContract.Words.Last
    wrdLast.Select
    
    Do While Selection.Characters.Count = 1
        ingWordLast = ingWordLast - 1
        Set wrdState = rngFirstParaOfContract.Words(ingWordLast) 'Move back a word at a time
        strState = wrdState.Text
        wrdState.Select
    Loop
    
    aResult = Filter(aCommonwealths, strState, True, vbTextCompare)
    If UBound(aResult) > -1 Then
        IsCommonWealth = True
        
        Set rngSelectedStateOfPhrase = Selection.Range
        rngSelectedStateOfPhrase.SetRange _
            Start:=rngFirstParaOfContract.Words(ingWordLast - 2).Start, _
            End:=wrdState.End
            
        rngSelectedStateOfPhrase.Select
    End If
    
    'Replace "State" with "Commonwealth"
    If IsCommonWealth Then
        Selection.Text = Replace(rngSelectedStateOfPhrase.Text, TARGET_TEXT_STATE, REPLACE_TEXT_COMMONWEALTH)
    End If
        
Exit_Here:
    Exit Sub

ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatCommonwealth"
    End With
    Resume Exit_Here
End Sub

Public Sub FormatCommonwealth_v2()
    '03/13/2018 I may retire these constants in favor of a search for "State of [STATE]
    Const PARA_COMMONWEALTH As Integer = 3
    
    '03/13/2018
    Const TARGET_TEXT_STATE As String = "State"
    Const REPLACE_TEXT_COMMONWEALTH As String = "Commonwealth"
    
    '03/22/2018
    Const strREGEX_COMMONWEALTH As String = "State of <*>"
    Const intWORD_COUNT_STATE As Integer = 3

    
    '03/12/2018
    Dim ingWordLast As Integer
    Dim wrdLast As Range
    Dim rngReplaceStateOfState As Range
    
    Dim rngFirstParaOfContract As Range
    Dim wrdState As Range
    Dim rngSelectedStateOfPhrase As Range
    Dim strState As String
    Dim IsCommonWealth As Boolean
    Dim aCommonwealths() As Variant
    Dim aResult() As String
    
    On Error GoTo ErrorHandler
    
    Selection.HomeKey Unit:=wdStory
    IsCommonWealth = False
    aCommonwealths = Array("Kentucky", "Massachusetts", "Pennsylvania", "Virginia")
    
    '03/22/2018 Addition
    Selection.Find.Execute FindText:=strREGEX_COMMONWEALTH, MatchWildcards:=True, MatchCase:=False
    If Selection.Find.Found Then
        Dim i As Integer
        Dim state As Variant
        For Each state In aCommonwealths
            If state = Selection.Words(intWORD_COUNT_STATE) Then
                Selection.Text = Replace(Selection.Text, TARGET_TEXT_STATE, REPLACE_TEXT_COMMONWEALTH)
                Exit For
            End If
        Next
    End If
        
Exit_Here:
    Exit Sub

ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatCommonwealth"
    End With
    Resume Exit_Here
End Sub

Public Sub FormatDateDayOfMonthYear()
'May 10, 2018
'Adrian Jones
    Dim DayDigit As Integer
    Dim DayAsOrdinal As String
    
    On Error GoTo ErrorHandler
    
    DayDigit = Val(Format(Now(), "dd"))
    DayAsOrdinal = FirstSecondThirdForth(CStr(DayDigit))
    Selection.Text = DayAsOrdinal & " of " & Format(Now(), "MMMM YYYY")
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Format Date: Day Of Month Year"
    End With
    Resume Exit_Here
End Sub
