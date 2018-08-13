'------------------------------------------------------------
' Command213_Click
'
'------------------------------------------------------------
Private Sub Command213_Click()
On Error GoTo Command213_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acNewRec
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If


Command213_Click_Exit:
    Exit Sub

Command213_Click_Err:
    MsgBox Error$
    Resume Command213_Click_Exit

End Sub
