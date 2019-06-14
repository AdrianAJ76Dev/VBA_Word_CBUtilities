VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAmendment 
   Caption         =   "Make Amendments"
   ClientHeight    =   2955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4665
   OleObjectBlob   =   "frmAmendment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAmendment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()
    On Error GoTo ErrorHandler
    
    Me.hide
    
Exit_Here:
    Exit Sub

ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "cmdOk_Click"
    End With
    Resume Exit_Here
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo ErrorHandler
    
    Me.Tag = vbNullString
    Me.hide
    
Exit_Here:
    Exit Sub

ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "optHED_Click"
    End With
    Resume Exit_Here
End Sub

Private Sub optHED_Click()
    On Error GoTo ErrorHandler
    
    Me.Tag = "HED"
    
Exit_Here:
    Exit Sub

ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "optHED_Click"
    End With
    Resume Exit_Here

End Sub

Private Sub optK12_Click()
    On Error GoTo ErrorHandler
    
    Me.Tag = "K12"
    
Exit_Here:
    Exit Sub

ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "optK12_Click"
    End With
    Resume Exit_Here
End Sub
