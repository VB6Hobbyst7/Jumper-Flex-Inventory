Private Sub add_G4_label_Click()
Dim amount As Variant
Dim SQL As String


' Get amount of G4 to add
amount = InputBox("How Many Parts Have Been Made?")
If amount <> vbNullString Then
    MsgBox "Adding " & amount & " elements"
    
    ' Add new G4 Elements
    SQL = "INSERT INTO G4_Test (Parts) VALUES (" & amount & ")"
    DoCmd.RunSQL SQL
Else
    MsgBox "No Elements Added"

End If
 
            
End Sub
