Attribute VB_Name = "protector"
Option Explicit

Public Const color_user_input = vbYellow

Const max_col As Integer = 20 ' as maximum allowed column to check

Dim the_row As Long, cells_in_row As Range, cell As Range


Function protection_handler(ws As Worksheet, rng As Range)
    
    On Error Resume Next
    the_row = rng.Row
    Set cells_in_row = Range(ws.Cells(the_row, 1), ws.Cells(the_row, max_col))
    tmp_unprotect ws
    For Each cell In cells_in_row
        If cell.Interior.Color = color_user_input Or _
           cell.DisplayFormat.Interior.Color = color_user_input Then
            If cell.Locked <> False Then
                cell.Locked = False
                cell.FormulaHidden = False
                cell.Calculate
            End If
        Else
            If cell.Locked <> True Then
                cell.Locked = True
                cell.FormulaHidden = True
                cell.Calculate
            End If
        End If
    Next
    
    tmp_protect ws

End Function

Private Function tmp_unprotect(ws As Worksheet)
    ws.Parent.Unprotect
    ws.Unprotect
End Function

Private Function tmp_protect(ws As Worksheet)
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    ws.Parent.Protect
End Function




