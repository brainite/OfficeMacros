VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RenameStyles 
   Caption         =   "Rename Word Styles"
   ClientHeight    =   2190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4485
   OleObjectBlob   =   "RenameStyles.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RenameStyles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub doCancel_Click()
    Me.Hide
End Sub

Private Sub doRenameStyles_Click()
    If Len(Me.TextMatch.Value) = 0 Then
        MsgBox "You must provide a Find value"
    End If

    Dim s As Style, success As Boolean
    On Error GoTo ErrHandler
    For Each s In ActiveDocument.Styles
        success = False
        s = Replace(s, Me.TextMatch.Value, Me.TextReplace.Value)
        success = True
        
ErrHandler:
        If Not success Then
            MsgBox "Unable to rename " & s
        End If
    Next s
    
    MsgBox "Finished!"
End Sub
