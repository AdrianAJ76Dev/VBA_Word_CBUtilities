Attribute VB_Name = "modRibbonInterface"
Option Explicit

Public Sub InterfaceCreateCoversheet()
    Dim clsCoversheet As clsCTmplateCoversheet
    On Error GoTo ErrorHandler
    
    Set clsCoversheet = New clsCTmplateCoversheet
    Documents.Add Template:=clsCoversheet.CoversheetTemplateName
    Application.ScreenRefresh
    clsCoversheet.CreateCoversheetForSignatureForSpecificContract InputBox("Enter or copy and past contract number", "Contract Number")

Exit_Here:
    Set clsCoversheet = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "InterfaceCreateCoversheet"
    End With
    Resume Exit_Here
End Sub

Public Sub InterfaceForSpellNumber()
    On Error GoTo ErrorHandler

    If IsNumeric(Selection.Text) Then
        Selection.Text = SpellNumber(Selection.Text)
    Else
        MsgBox "This only works on numbers.  Please select a numbers to spell out", vbInformation, "Spell Out Number"
    End If

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "InterfaceForSpellNumber"
    End With
    Resume Exit_Here
End Sub

Public Sub InterfaceForTwoWeeksFromToday()
    On Error GoTo ErrorHandler
    
    Selection.Text = TwoWeeksFromToday

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "InterfaceForTwoWeeksFromToday"
    End With
    Resume Exit_Here
End Sub
