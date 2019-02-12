Private Sub btnNewUpdate_Click()

Dim SQL As String
Dim dbs As DAO.Database
Dim amount As Variant
Dim rsStock As DAO.Recordset
Dim name_str As String
Dim test As Integer


[lblName].Visible = True
[btnNewUpdate].Visible = True
[txtboxName].Visible = True


Set dbs = CurrentDb
Set rsStock = dbs.OpenRecordset("CTF_open_orders", dbOpenSnapshot)
name_str = [Forms]![open_orders_form]![txtboxName]


' dbs.Execute "UPDATE CTF_open_orders " _
' & "SET Statuts = 'Done' " _
' & "WHERE Customer = '" & name_str & "';"


'amount = InputBox(Prompt:="Enter New Status " & vbCrLf & " (Hit Ok on the next screen)", Title:="Change Status")
If txtboxName <> vbNullString Then
    ' modify the sql statement to update if the customer name is not-valid
    ' SQL is working, but not in the way we want it to
    ' A random Msbox displayd during [...].Customer phase of the program
    ' Seems to have a problem with spaces in the Customer Name
    SQL = "UPDATE [CTF_open_orders] SET [CTF_open_orders].Status = 'DONE' WHERE [CTF_open_orders].Customer = '" & name_str & "';"
    DoCmd.RunSQL SQL
Else
    ' modify the if statement
       MsgBox " & name_str & " & vbNewLine & "Click Yes to Continue", vbYesNo + vbQuestion, "Customer Not Found"

End If


End Sub
