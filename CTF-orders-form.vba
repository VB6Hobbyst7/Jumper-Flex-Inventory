Public whatsClicked As String
Option Compare Database

Private Sub btnChangeOrderNotes_Click()
' Update the notes field in the CTF OPEN ORDERS Table

' set the flag
whatsClicked = "change-notes"

[txtboxName].Visible = True
[txtNewNotes].Visible = True
[lblName].Visible = True
[lblNewNotes].Visible = True
[btnNewUpdate].Visible = True
[btnOpenFolder].Visible = False

End Sub

Private Sub btnChangeStatus_Click()
Dim txtCustomer As Variant
Dim SQL As String
Dim done_str As String
Dim dbs As DAO.Database
Dim rsStock As DAO.Recordset

' set the flag
whatsClicked = "change-status"

[lblName].Visible = True
[lblStatus].Visible = True
[btnNewUpdate].Visible = True
[txtboxName].Visible = True
[txtStatus].Visible = True
[lblNewNotes].Visible = False
[btnNewElementOrder].Visible = False
[btnNewFixtureOrder].Visible = False
[arrowImg].Visible = False
[btnChangeOrderNotes].Visible = False
[btnOpenFolder].Visible = False

            
            
End Sub

Private Sub btnClose_Click()
' Hide the finished Orders Query and re-open the Open Orders

[CTF_open_orders_Query_subform].Visible = True
[CTF_open_orders_ALL_ORDERS_query subform].Visible = False
[btnHome].SetFocus
[btnClose].Visible = False
[lblStatus].Visible = False
[lblName].Visible = False
[btnNewUpdate].Visible = False
[txtboxName].Visible = False
[txtStatus].Visible = False
[btnNewElementOrder].Visible = False
[btnNewFixtureOrder].Visible = False
[arrowImg].Visible = False
[btnChangeOrderNotes].Visible = True
[btnOpenFolder].Visible = False


End Sub

Private Sub btnCustomerFolder_Click()
' open up the customer folders in 'sales and marketing'
Dim Customer As String

[lblName].Visible = True
[txtboxName].Visible = True
[btnOpenFolder].Visible = True


End Sub

Private Sub btnHome_Click()

[lblName].Visible = False
[lblStatus].Visible = False
[btnNewUpdate].Visible = False
[txtboxName].Visible = False
[txtStatus].Visible = False
[btnNewElementOrder].Visible = False
[btnNewFixtureOrder].Visible = False
[arrowImg].Visible = False
[txtNewNotes].Visible = False
[btnOpenFolder].Visible = False

DoCmd.Close
DoCmd.OpenForm "Home Page"

End Sub

Private Sub btnNewElementOrder_Click()

DoCmd.Close
DoCmd.OpenForm "Add_Element_Order_Form"
' [btnOpenFolder].Visible = False

End Sub

Private Sub btnNewFixtureOrder_Click()


DoCmd.Close
DoCmd.OpenForm "Add_CTF_Order"
' [btnOpenFolder].Visible = False


End Sub

Private Sub btnNewUpdate_Click()

Dim SQL As String
Dim dbs As DAO.Database
Dim amount As Variant
Dim rsStock As DAO.Recordset
Dim name_str As String
Dim status_str As String
Dim test As Integer

DoCmd.SetWarnings False

[lblName].Visible = True
[lblStatus].Visible = True
[btnNewUpdate].Visible = True
[txtboxName].Visible = True
[txtStatus].Visible = True
[btnNewElementOrder].Visible = False
[btnNewFixtureOrder].Visible = False
[arrowImg].Visible = False
[btnOpenFolder].Visible = False

Set dbs = CurrentDb
Set rsStock = dbs.OpenRecordset("CTF_open_orders", dbOpenSnapshot)
name_str = [Forms]![open_orders_form]![txtboxName]


    ' Find out what button is selected
    Select Case whatsClicked
    Case "change-status":
        status_str = [Forms]![open_orders_form]![txtStatus]
        SQL = "UPDATE [CTF_open_orders] SET [CTF_open_orders].Status = '" & status_str & "' WHERE [CTF_open_orders].Status = 'Open' AND [CTF_open_orders].Customer = '" & name_str & "';"
        DoCmd.RunSQL SQL
        ' Confirmation MsgBox
        MsgBox name_str & " status changed to " & status_str
    Case "change-notes":
        status_str = [Forms]![open_orders_form]![txtNewNotes]
        SQL = "UPDATE [CTF_open_orders] SET [CTF_open_orders].Notes = '" & status_str & "' WHERE [CTF_open_orders].Customer = '" & name_str & "';"
        DoCmd.RunSQL SQL
        MsgBox name_str & " notes changed to " & status_str
    End Select
    
    
' clear the text boxes
' getting an error message here bc the textboxes don't have the focus
' [txtboxName].Text = " "
' [txtStatus].Text = " "
' [txtNewNotes].Text = " "
    
    
[CTF_open_orders_ALL_ORDERS_query subform].Visible = False
[btnClose].Visible = False
[lblName].Visible = False
[lblStatus].Visible = False

[txtboxName].Visible = False
[txtStatus].Visible = False
[txtNewNotes].Visible = False
[lblNewNotes].Visible = False
[btnNewElementOrder].Visible = False
[btnNewFixtureOrder].Visible = False
[arrowImg].Visible = False
[btnOpenFolder].Visible = False


End Sub

Private Sub btnOpenFolder_Click()
' open the selected customer folder
Dim Customer As String
Dim full_path As String

Customer = [Forms]![open_orders_form]![txtboxName]
full_path = "\\zdata\ZData\Sales and Marketing\2019\customers\" & Customer

Call Shell("explorer.exe" & " " & "\\zdata\ZData\Sales and Marketing\2019\customers\" & Customer, vbNormalFocus)

End Sub

Private Sub Command21_Click()

' may be easier to open up a differnet form
'get_customer_name = InputBox(Prompt:="Enter Customer Name " & vbCrLf & " (Hit Ok on the next screen)", Title:="Add New Order")

' DoCmd.OpenForm "Add_CTF_Order" This is form the fixtures
[btnNewElementOrder].Visible = True
[btnNewFixtureOrder].Visible = True
[arrowImg].Visible = True
[btnOpenFolder].Visible = False

End Sub

Private Sub Command3_Click()

[CTF_open_orders_Query_subform].Visible = False
[CTF_open_orders_ALL_ORDERS_query subform].Visible = True
[btnClose].Visible = True
[lblName].Visible = False
[lblStatus].Visible = False
[btnNewUpdate].Visible = False
[btnChangeOrderNotes].Visible = False
[txtboxName].Visible = False
[txtStatus].Visible = False
[txtNewNotes].Visible = False
[btnNewElementOrder].Visible = False
[btnNewFixtureOrder].Visible = False
[arrowImg].Visible = False
[btnOpenFolder].Visible = False
' Add another cmd to revert back to OPEN ORDERS

End Sub

Private Sub Form_Load()
' when the form is open, ALL ORDERS QUERY will be hidden

[CTF_open_orders_ALL_ORDERS_query subform].Visible = False
[btnClose].Visible = False
[lblName].Visible = False
[lblStatus].Visible = False
[btnNewUpdate].Visible = False
[txtboxName].Visible = False
[txtStatus].Visible = False
[txtNewNotes].Visible = False
[lblNewNotes].Visible = False
[btnNewElementOrder].Visible = False
[btnNewFixtureOrder].Visible = False
[arrowImg].Visible = False
[btnOpenFolder].Visible = False

End Sub
