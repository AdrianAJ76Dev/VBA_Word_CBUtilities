VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbonEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjWordApp As Word.Application
Attribute mobjWordApp.VB_VarHelpID = -1
Private mobjRibbonUI As IRibbonUI

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set mobjWordApp = Application

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Raise .Number & .Description, "clsRibbonEvents:Class_Initialize"
    End With
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Set mobjWordApp = Nothing
    Set mobjRibbonUI = Nothing
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Raise .Number & .Description, "clsRibbonEvents:Class_Initialize"
    End With
End Sub

Public Property Set Ribbon(ByRef robjribbon As IRibbonUI)
    On Error GoTo ErrorHandler
    
    Set mobjRibbonUI = robjribbon
    
Exit_Here:
    Exit Property
    
ErrorHandler:
    With Err
        .Raise .Number & .Description, "clsRibbonEvents:Class_Initialize"
    End With
End Property

Private Sub mobjWordApp_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
    On Error GoTo ErrorHandler

    mobjRibbonUI.Invalidate

Exit_Here:
    Exit Sub

ErrorHandler:
    With Err
        .Raise .Number & .Description, "clsRibbonEvents:mobjWordApp_DocumentBeforeClose"
    End With
End Sub

Private Sub mobjWordApp_DocumentOpen(ByVal Doc As Document)
    On Error GoTo ErrorHandler

    mobjRibbonUI.Invalidate

Exit_Here:
    Exit Sub

ErrorHandler:
    With Err
        .Raise .Number & .Description, "clsRibbonEvents:mobjWordApp_DocumentOpen"
    End With
End Sub

Private Sub mobjWordApp_NewDocument(ByVal Doc As Document)
    On Error GoTo ErrorHandler

    mobjRibbonUI.Invalidate

Exit_Here:
    Exit Sub

ErrorHandler:
    With Err
        .Raise .Number & .Description, "clsRibbonEvents:mobjWordApp_NewDocument"
    End With
End Sub

Private Sub mobjWordApp_Quit()
    On Error GoTo ErrorHandler
    
    Set mobjWordApp = Nothing
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        .Raise .Number & .Description, "clsRibbonEvents:mobjWordApp_Quit"
    End With
End Sub
