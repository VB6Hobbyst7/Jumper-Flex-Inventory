Private Sub add_G4_label_Click()
Dim amount As Variant
Dim sql As String


' Get amount of G4 to add, also add the SQL statement to save the data with new IDs
amount = InputBox("How Many Parts Have Been Made?")
If amount <> vbNullString Then
    MsgBox "Adding " & amount & " elements"
    
    ' Add new G4 Elements, get rid of 'TEST' in location
    sql = "INSERT INTO G4 Test VALUES ('" & amount & "')"
    
    'DoCmd.RunSQL sql
    'DoCmd.RunSQL "INSERT INTO G4 Test VALUES ('" & amount & "')"
Else
    MsgBox "No Elements Added"

End If
 
            
End Sub
