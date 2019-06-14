Attribute VB_Name = "modRevisions"
Option Explicit

Public Sub GetRevisions()
    Dim rev As Revision
    On Error GoTo ErrorHandler
    
    
    For Each rev In ActiveDocument.Revisions
        If rev.Type = wdRevisionInsert Then
            Debug.Print rev.Range.Text
        End If
    Next rev

ErrorHandler:
    With Err
        MsgBox .Number & vbCr & .Description, vbCritical, "GetRevisions"
    End With
End Sub
