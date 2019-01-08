Private Sub add_G4_label_Click()
Dim amount As Variant
Dim SQL As String

' Get amount of G4 to add
amount = InputBox(Prompt:="How many blanks have been made? " & vbCrLf & " (Hit Ok on the next screen)", Title:="Add G4 Stock")
' modify the Or statement, not working
If amount <> vbNullString And amount <> 0 Then
    SQL = "UPDATE [G4_Stock] SET [G4_Stock].Full_blanks = [G4_Stock]![Full_Blanks] + " & amount & ";"
    DoCmd.RunSQL SQL
Else
    MsgBox "No Elements Added"

End If
             
End Sub
