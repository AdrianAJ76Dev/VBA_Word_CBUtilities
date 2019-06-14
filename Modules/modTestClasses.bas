Attribute VB_Name = "modTestClasses"
Option Explicit
Private mrbnEvt As clsRibbonEvents

Public Sub TestRiderClass()
    Dim CBRiders As clsCBRiders
    On Error GoTo ErrorHandler
    
    Set CBRiders = New clsCBRiders
    CBRiders.CleanUpRiders
    
Exit_Here:
    Set CBRiders = Nothing
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCrLf & .Description, vbCritical, "TestRiderClass"
    End With
    Resume Exit_Here
End Sub

Public Sub TestRibbonEventsOn()
    On Error GoTo ErrorHandler
    
    'This reliably sets up the capture of Application events of course until the application, Word in this case, Closes.
    Set mrbnEvt = New clsRibbonEvents
'    Set mrbnEvt.Ribbon = RibbonUI
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCrLf & .Description, vbCritical, "TestRibbonEventsOn"
    End With
    Resume Exit_Here
End Sub

Public Sub TestRibbonEventsOff()
    
    On Error GoTo ErrorHandler
    
    'This reliably sets up the capture of Application events of course until the application, Word in this case, Closes.
    Set mrbnEvt = Nothing
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCrLf & .Description, vbCritical, "TestRibbonEventsOff"
    End With
    Resume Exit_Here
End Sub

