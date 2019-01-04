Private Sub label_G4_Click()
' Display how many blanks in stock

Dim total_blanks As Integer
Dim dbs As DAO.Database
Dim rsStock As DAO.Recordset

Set dbs = CurrentDb
Set rsStock = dbs.OpenRecordset("G4_Stock", dbOpenSnapshot)

total_blanks = DSum("Full_Blanks", "G4_Stock")

MsgBox total_blanks & " G4 blanks availabe" & vbCrLf & "( In Bin 1 )"

End Sub
