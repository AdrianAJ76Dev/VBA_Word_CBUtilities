Attribute VB_Name = "modSampleCode"
Option Explicit

Public Sub DocumentNewTest()
    On Error GoTo ErrorHandler
    
    Application.Documents.Add Template:="C:\Documents and Settings\ajones\Application Data\Microsoft\Templates\Contracts Management\College Board GDPR Data Sharing Addendum.dotx"

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "DocumentNewTest"
    End With
    Resume Exit_Here
End Sub

Public Sub TestClass()
    Dim objClassTested As clsCBContract
    On Error GoTo ErrorHandler
    
    '*********************************************************************************************************
    Set objClassTested = New clsCBContract 'Change here to test any created class
'    objClassTested.DetermineActiveRiders
    objClassTested.ActivateRider
    '*********************************************************************************************************

Exit_Here:
    Set objClassTested = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "TestClass"
    End With
    Resume Exit_Here
End Sub

Sub Insert_Updated_Dates()
Attribute Insert_Updated_Dates.VB_Description = "Inserts updated, correct dates for the fiscal year into the riders"
Attribute Insert_Updated_Dates.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Insert_Updated_Dates"
'
' Insert_Updated_Dates Macro
' Inserts updated, correct dates for the fiscal year into the riders
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "September 14, 2013"
        .Replacement.Text = "September 13, 2013"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "September 14, 2013"
        .Replacement.Text = "September 13, 2013"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "September 3, 2013"
        .Replacement.Text = "September 2, 2013"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub Replace2012()
Attribute Replace2012.VB_Description = "Search and Replace 2012-2013 with 2013-2014"
Attribute Replace2012.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Replace2012"
'
' Replace2012 Macro
' Search and Replace 2012-2013 with 2013-2014
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub TableFormat()
Attribute TableFormat.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.TableFormat"
'
' TableFormat Macro
'
'
    Selection.Tables(1).Rows.Alignment = wdAlignRowLeft
    Selection.Tables(1).Rows.WrapAroundText = False
End Sub
