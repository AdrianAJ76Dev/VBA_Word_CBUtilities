VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDocumentVariables 
   Caption         =   "Document Variables"
   ClientHeight    =   4764
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   3780
   OleObjectBlob   =   "frmDocumentVariables.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDocumentVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdClose_Click()
    On Error GoTo ErrorHandler

    Unload Me

Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "cmdClose_Click"
    End With
    Resume Exit_Here
End Sub

Private Sub UserForm_Initialize()
    Dim objDocVar As Variable
    Dim i As Integer
    On Error GoTo ErrorHandler
    
    i = -1
    For Each objDocVar In ActiveDocument.Variables
        Me.lstDocumentVariables.AddItem objDocVar.Name, i
        i = i + 1
    Next objDocVar
    
Exit_Here:
    Exit Sub
    
ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "UserForm Initialize"
    End With
    Resume Exit_Here
End Sub
