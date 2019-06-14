Attribute VB_Name = "modRecordedMacros"
Option Explicit

Sub MoveToTheTop()
Attribute MoveToTheTop.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.MoveToTheTop"
'
' MoveToTheTop Macro
'
'
    Selection.HomeKey Unit:=wdStory
    Selection.MoveDown Unit:=wdScreen, Count:=35, Extend:=wdExtend
    Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
End Sub
Sub GotoPage()
Attribute GotoPage.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.GotoPage"
'
' GotoPage Macro
'
'
    Selection.HomeKey Unit:=wdStory
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Name:="8"
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ""
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
    Selection.HomeKey Unit:=wdStory
    Selection.Extend
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Name:="8"
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ""
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
End Sub
Sub RemoveSurroundingTablesRecordedMacro()
Attribute RemoveSurroundingTablesRecordedMacro.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.RemoveSurroundingTables"
'
' RemoveSurroundingTables Macro
'
'
    Selection.Rows.ConvertToText Separator:=wdSeparateByParagraphs, _
        NestedTables:=False
End Sub
