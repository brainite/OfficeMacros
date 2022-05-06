VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormAlignRows 
   Caption         =   "FormAlignRows"
   ClientHeight    =   2865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4260
   OleObjectBlob   =   "FormAlignRows.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormAlignRows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LeftCol_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)

End Sub

Private Sub RealignRows_Click()
    Dim r As Long, leftColNum As Long, lastRow As Long, lastCol As Long, d As Range, selection As Range
    Dim strShift As String, strLeft As String, strRight As String
    Dim lngI As Long, lngFoundLeft As Long, lngFoundRight As Long
    Dim blnContinue As Boolean, intResponse As Integer
    
    Set selection = Range(LeftCol.Value)
    ' Application.ScreenUpdating = False
    lastRow = selection.Worksheet.UsedRange.Rows.Count
    lastCol = selection.Worksheet.UsedRange.Columns.Count
    
    Set d = Range("A1:A" & (lastRow + 200))
    
    ' Set "r" to the first row
    r = selection.Row
    leftColNum = selection.Column
    
    
    ' Iterate until the row's value is empty
    blnContinue = True
    Do While blnContinue
        strLeft = d.Cells(r, leftColNum)
        strRight = d.Cells(r, leftColNum).Offset(, 1)
            
        ' If RIGHT is not empty...
        If strRight <> "" And strLeft <> "" Then
            ' No shift needed; LEFT = RIGHT
            If strLeft = strRight Then
                ' DO NOTHING.
            Else
                ' See whether the refs exist in the next X rows
                lngFoundLeft = 0
                lngFoundRight = 0
                lngI = 0
                Do While lngI < 20
                    If strRight <> "" And lngFoundRight = 0 And d.Cells(r + lngI, leftColNum) = strRight Then
                        lngFoundRight = r + lngI
                    End If
                    If strLeft <> "" And lngFoundLeft = 0 And d.Cells(r + lngI, leftColNum).Offset(, 1) = strLeft Then
                        lngFoundLeft = r + lngI
                    End If
                    lngI = lngI + 1
                Loop
                
                ' Adjust the shift.
                strShift = "R"
                If lngFoundLeft = 0 And lngFoundRight = 0 Then
                    If strLeft > strRight Then
                        strShift = "L"
                    End If
                ' If LEFT<RIGHT, shift LEFT down
                ' If LEFT=0, RIGHT>0, shift RIGHT down
                ' If LEFT>RIGHT, shift RIGHT down
                ' If LEFT>0, RIGHT=0, shift LEFT down
                ElseIf (lngFoundLeft > 0 And lngFoundRight = 0) Or (lngFoundLeft > 0 And lngFoundLeft < lngFoundRight) Then
                    strShift = "L"
                End If
            
                ' Insert cells as appropriate
                If strShift = "R" Then
                    Application.StatusBar = "Shift right with cols=" & (lastCol - leftColNum)
                    d.Cells(r, leftColNum).Offset(, 1).Resize(, lastCol - leftColNum).Insert xlShiftDown
                Else
                    d.Cells(r, 1).Resize(, leftColNum).Insert xlShiftDown
                    lastRow = lastRow + 1
                    Set d = Range("A1:A" & lastRow)
                End If
            End If
        End If
      
        ' Look at the next row to decide whether to continue.
        r = r + 1
        Application.StatusBar = "Inspecting row #" & r
        blnContinue = (d.Cells(r, leftColNum) <> "")
        If r Mod 100 = 0 Then
            intResponse = MsgBox("Looking at row #" & r & ". Continue?", vbYesNo + vbQuestion)
            If vbNo = intResponse Then
                blnContinue = False
            End If
        End If
    Loop
    Application.ScreenUpdating = True
    Me.Hide
End Sub
