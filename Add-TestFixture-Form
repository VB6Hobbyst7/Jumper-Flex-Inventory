Private Sub btnAddOrder_Click()
' Add a New CTF order and check if the correct flexes are in stock.
' If no flexes in stock, notify user to order flexes

Dim dbs As DAO.Database
Dim amount As Variant
Dim rsStock As DAO.Recordset
Dim rsFlex As DAO.Recordset

Dim txt_name As String
Dim txt_pitch As Double
Dim txt_contacts As Integer
Dim txt_quantity As Integer
Dim flexSQL As String
Dim txt_partnumber As String
Dim txt_flexlength As String    ' may have to use an int for this
Dim txt_status As String
Dim txt_shipby As String
Dim flexQuery As DAO.Recordset
' Var to check the flexes in stock
Dim flexes_in_stock As Variant

DoCmd.SetWarnings False

Set dbs = CurrentDb
Set rsStock = dbs.OpenRecordset("CTF_open_orders", dbOpenSnapshot)
Set rsFlex = dbs.OpenRecordset("Copy Of Jumper Flex Inventory--PRODUCTION", dbOpenSnapshot)


txt_name = [Forms]![Add_CTF_Order]![nameBox]
txt_pitch = [Forms]![Add_CTF_Order]![pitchBox]
txt_contacts = [Forms]![Add_CTF_Order]![ContactsBox]
txt_quantity = [Forms]![Add_CTF_Order]![quantityBox]
txt_partnumber = [Forms]![Add_CTF_Order]![numberBox]
txt_status = [Forms]![Add_CTF_Order]![statusBox]
txt_shipby = [Forms]![Add_CTF_Order]![shipBox]
txt_flexlength = [Forms]![Add_CTF_Order]![txtFLexLength]


' Running a SQL to check for flexes in stock -- for now
' flexSQL = "SELECT [Copy Of Jumper Flex Inventory--PRODUCTION].Pitch_mm, [Copy Of Jumper Flex Inventory--PRODUCTION].[Number of Conductors], [Copy Of Jumper Flex Inventory--PRODUCTION].Length_in FROM [Copy Of Jumper Flex Inventory--PRODUCTION] WHERE ((([Copy Of Jumper Flex Inventory--PRODUCTION].Pitch_mm)=  '" & txt_pitch & "' And (([Copy Of Jumper Flex Inventory--PRODUCTION].[Number of Conductors])= '" & txt_contacts & "') And (([Copy Of Jumper Flex Inventory--PRODUCTION].Length_in)= '" & txt_flexlength & "'));"

DoCmd.RunSQL "INSERT INTO CTF_Open_Orders (Customer, Pitch, Contacts, Quantity, PartNumber, Status, ShipDate) VALUES ('" & txt_name & "', '" & txt_pitch & "', '" & txt_contacts & "', '" & txt_quantity & "', '" & txt_partnumber & "', '" & txt_status & "', '" & txt_shipby & "')"
' MsgBox txt_name & " added to orders"

' Test checking the flex quantity
' DoCmd.RunSQL flexSQL
MsgBox "The flex length is: " + txt_flexlength

Form.Refresh

' count if there are any flexes returned by the query
' somewhat works but if there is nothing there -- it'll just return null
flexes_in_stock = DSum("Quantity", "checkFlexesAfterOrder_Query")

' if the the quantity value is 0, display a MsgBox
If IsNull(flexes_in_stock) Then
    MsgBox "There are no flexes in stock"
Else
    lblCheckFlex.Visible = True
    Label97.Visible = True
    checkFlexesAfterOrder_QuerySubform.Visible = True
End If


End Sub
