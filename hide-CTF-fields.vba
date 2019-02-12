Option Compare Database

Private Sub btnChangeStatus_Click()

Dim txtCustomer As Variant
Dim SQL As String
Dim done_str As String
Dim dbs As DAO.Database
Dim rsStock As DAO.Recordset

[lblName].Visible = True
[btnNewUpdate].Visible = True
[txtboxName].Visible = True
           
End Sub


Private Sub btnClose_Click()
' Hide the finished Orders Query and re-open the Open Orders

[CTF_open_orders_Query_subform].Visible = True
[CTF_open_orders_ALL_ORDERS_query subform].Visible = False
[btnHome].SetFocus
[btnClose].Visible = False
[lblName].Visible = False
[btnNewUpdate].Visible = False
[txtboxName].Visible = False

End Sub


Private Sub btnHome_Click()

[lblName].Visible = False
[btnNewUpdate].Visible = False
[txtboxName].Visible = False
DoCmd.Close
DoCmd.OpenForm "APP-START"

End Sub


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

Private Sub Command3_Click()

[CTF_open_orders_Query_subform].Visible = False
[CTF_open_orders_ALL_ORDERS_query subform].Visible = True
[btnClose].Visible = True
[lblName].Visible = False
[btnNewUpdate].Visible = False
[txtboxName].Visible = False
' Add another cmd to revert back to OPEN ORDERS

End Sub

Private Sub Form_Load()
' when the form is open, ALL ORDERS QUERY will be hidden

[CTF_open_orders_ALL_ORDERS_query subform].Visible = False
[btnClose].Visible = False
[lblName].Visible = False
[btnNewUpdate].Visible = False
[txtboxName].Visible = False

End Sub
