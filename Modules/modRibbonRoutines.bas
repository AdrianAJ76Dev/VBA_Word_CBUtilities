Attribute VB_Name = "modRibbonRoutines"
Option Explicit
Private mobjRibbonEvents As clsRibbonEvents
'********************************************************************************************************************************************************************************************
'Create Date:   04/27/2017
'Creating routines to couple shortcut keys to ribbon buttons
'********************************************************************************************************************************************************************************************
'Stub
Public Sub ROUTINE_NAME_Ribbon(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    'This is the stub code that should be copied when creating a ribbon button to run code

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "ROUTINE_NAME_Ribbon"
    End With
    Resume Exit_Here
End Sub
'Stub
'********************************************************************************************************************************************************************************************

Public Sub RibbonEventsInialized_Ribbon(ByVal RibbonUI As IRibbonUI)
    On Error GoTo ErrorHandler
    
    Set mobjRibbonEvents = New clsRibbonEvents
    Set mobjRibbonEvents.Ribbon = RibbonUI
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "RibbonEventsInialized_Ribbon"
    End With
    Resume Exit_Here
End Sub

Public Sub RibbonEventsTerminated_Ribbon(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Set mobjRibbonEvents = Nothing
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "RibbonEventsTerminated_Ribbon"
    End With
    Resume Exit_Here
End Sub

Public Sub Clean_Up_Riders_Ribbon(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Clean_Up_Riders

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "Clean_Up_Riders_Ribbon"
    End With
    Resume Exit_Here
End Sub

Public Sub RefreshShortcuts_Ribbon(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    RefreshShortcuts

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "RefreshShortcuts_Ribbon"
    End With
    Resume Exit_Here
End Sub

Public Sub FormatPrice_Ribbon(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    FormatPrice

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatPrice_Ribbon"
    End With
    Resume Exit_Here
End Sub

Public Sub FormatDateSpellOutMonth_Ribbon(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    FormatDateSpellOutMonth

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatDateSpellOutMonth_Ribbon"
    End With
    Resume Exit_Here
End Sub

Public Sub FormatPhoneNumber_Ribbon(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    FormatPhoneNumber

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatPhoneNumber_Ribbon"
    End With
    Resume Exit_Here
End Sub

Public Sub FormatCommonwealth_Ribbon(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    FormatCommonwealth_v2

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "FormatCommonwealth_v2_Ribbon"
    End With
    Resume Exit_Here
End Sub


Public Sub CreateSoleSourceLetter_Ribbon(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    CreateSoleSourceLetter

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "CreateSoleSourceLetter_Ribbon"
    End With
    Resume Exit_Here
End Sub

Public Sub MakeHEDAmendment_Ribbon(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    MakeAmendment
     
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "MakeHEDAmendment_Ribbon"
    End With
    Resume Exit_Here
End Sub

Public Sub CreateCoversheetForSignature_Ribbon(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    InterfaceCreateCoversheet
        
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "CreateCoversheetForSignature_Ribbon"
    End With
    Resume Exit_Here
End Sub

Public Sub DeleteMyRoad_Ribbon(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    FindTextToDelete "MyRoad*and"

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "DeleteMyRoad_Ribbon"
    End With
    Resume Exit_Here
End Sub

Public Sub GetEnabled(ByVal control As IRibbonControl, ByRef rblnEnabled)
    On Error GoTo ErrorHandler
    
    'Always assume the best :-)
    rblnEnabled = True
    If Application.Documents.Count = 0 Then
        rblnEnabled = False
    ElseIf Application.Documents(1).Name = ThisDocument.Name Then
        rblnEnabled = False
    End If
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "GetEnabled"
    End With
    Resume Exit_Here
End Sub

Public Sub InterfaceForSpellNumber_Ribbon(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    InterfaceForSpellNumber

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "InterfaceForSpellNumber_Ribbon"
    End With
    Resume Exit_Here
End Sub

Public Sub InterfaceForTwoWeeksFromToday_Ribbon(ByVal control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    InterfaceForTwoWeeksFromToday

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "InterfaceForTwoWeeksFromToday_Ribbon"
    End With
    Resume Exit_Here
End Sub
