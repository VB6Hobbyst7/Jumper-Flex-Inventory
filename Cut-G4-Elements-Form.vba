' Clean this up
' Add feature to check if the element already exists, and if it does, just add on to the total
' Instead of a making a whole new row
Option Compare Database

' --------------------------------------------------
Private Sub btnAddG4_Click()

lblPitch.Visible = True
lblTraces.Visible = True
lblLocation.Visible = True
lblQuantity.Visible = True
txtPitch.Visible = True
txtTraces.Visible = True
txtLocation.Visible = True
txtQuantity.Visible = True
btnAddNewG4.Visible = True

End Sub
' --------------------------------------------------


' --------------------------------------------------
Private Sub btnAddNewG4_Click()
' add the new G4 elements to the table
Dim dbs As DAO.Database
Dim rsStock As DAO.Recordset
Dim Pitch As Double
Dim traces As Integer
Dim Quantity As Integer
Dim Location As String

' get values from textboxes
Pitch = [Forms]![find_g4_elements]![txtPitch]
traces = [Forms]![find_g4_elements]![txtTraces]
Quantity = [Forms]![find_g4_elements]![txtQuantity]
Location = [Forms]![find_g4_elements]![txtLocation]

Set dbs = CurrentDb
Set rsStock = dbs.OpenRecordset("Cut_G4_Elements", dbOpenSnapshot)

DoCmd.RunSQL "INSERT INTO Cut_G4_Elements (Pitch, Traces, Quantity, Location) VALUES ('" & Pitch & "', '" & traces & "', '" & Quantity & "', '" & Location & "')"
MsgBox Pitch & " x " & traces & " G4 element added to database"

' hide everything when complete
findG4btn.SetFocus
lblPitch.Visible = False
lblTraces.Visible = False
lblLocation.Visible = False
lblQuantity.Visible = False
txtQuantity.Visible = False
txtPitch.Visible = False
txtTraces.Visible = False
txtLocation.Visible = False
btnAddNewG4.Visible = False

End Sub
' --------------------------------------------------


' --------------------------------------------------
Private Sub btnUpdate_Click()
Dim SQL As String
Dim Quantity_Out As Integer
Dim test As Integer
Dim Contacts As Integer
Dim dbs As DAO.Database
Dim rsStock As DAO.Recordset

Set dbs = CurrentDb
Set rsStock = dbs.OpenRecordset("Cut_G4_Elements", dbOpenSnapshot)
Quantity_Out = [Forms]![find_g4_elements]![txtQuantityOut]
Contacts = [Forms]![find_g4_elements]![g4_traces_text]
test = [Forms]![Cut-G4-Elements-Query_subform_good]![Quantity]

' Update and use the QUERY instead of refrencing the table
' MODIFY THE STATEMENT
' Seems to be getting the right elements, ONLY when 'Cut-G4-Elements-Query_subform_good' is OPEN
' Add extra Docmd.Requery() to refresh the QUERY DATA
 SQL = "UPDATE [Cut_G4_Elements] SET [Cut_G4_Elements.Quantity] = [Cut_G4_Elements]![Quantity] - " & Quantity_Out & " WHERE [Cut_G4_Elements.Traces] = '" & Contacts & "';"
' DoCmd.RunSQL SQL

MsgBox test
End Sub
' --------------------------------------------------


' --------------------------------------------------
Private Sub cut_g4_save_Click(
' Close the form and go back to INDEX PAGE

DoCmd.Close
DoCmd.OpenForm "Home Page"

End Sub
' --------------------------------------------------


' --------------------------------------------------
Private Sub findG4btn_Click()
' Run the query and refresh the page
' add a delete option for elements

' hide the add elements
lblPitch.Visible = False
lblTraces.Visible = False
lblLocation.Visible = False
lblQuantity.Visible = False
txtQuantity.Visible = False
txtPitch.Visible = False
txtTraces.Visible = False
txtLocation.Visible = False
btnAddNewG4.Visible = False

DoCmd.Requery

' If Forms![Cut_G4_Elements_subform]![Traces] = Null Then
End Sub
' --------------------------------------------------


' --------------------------------------------------
Private Sub findG4btn_Enter()
' Run the query if the enter key is pressed
DoCmd.Requery
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
DoCmd.Requery
End Sub
' --------------------------------------------------


' --------------------------------------------------
Private Sub Form_Load()
' hide the 'add-g4' related stuff
lblPitch.Visible = False
lblTraces.Visible = False
lblLocation.Visible = False
lblQuantity.Visible = False
txtQuantity.Visible = False
txtPitch.Visible = False
txtTraces.Visible = False
txtLocation.Visible = False
btnAddNewG4.Visible = False

End Sub
' --------------------------------------------------
