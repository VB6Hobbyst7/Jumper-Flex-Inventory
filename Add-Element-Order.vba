Option Compare Database

Private Sub btnAddOrder_Click()
Dim dbs As DAO.Database
Dim rsStock As DAO.Recordset

Dim txt_name As String
Dim txt_pitch As Double
Dim txt_contacts As Integer
Dim txt_quantity As Integer
Dim txt_partnumber As String
Dim txt_status As String
Dim txt_shipby As String

DoCmd.SetWarnings False

' pitchBox.Text = 0.5
' numberBox.Text = "ZFLAT-CTF07-"

' assign the VARs here...
txt_name = [Forms]![Add_Element_Order_Form]![nameBox]
txt_pitch = [Forms]![Add_Element_Order_Form]![pitchBox]
txt_contacts = [Forms]![Add_Element_Order_Form]![ContactsBox]
txt_quantity = [Forms]![Add_Element_Order_Form]![quantityBox]
txt_partnumber = [Forms]![Add_Element_Order_Form]![numberBox]
txt_status = [Forms]![Add_Element_Order_Form]![statusBox]
txt_shipby = [Forms]![Add_Element_Order_Form]![shipBox]

If txt_pitch = "0.5" Then txt_name = txt_name & "--G4"
If txt_pitch = "1.0" Then txt_name = txt_name & "--C"

Set dbs = CurrentDb
Set rsStock = dbs.OpenRecordset("CTF_open_orders", dbOpenSnapshot)

DoCmd.RunSQL "INSERT INTO CTF_Open_Orders (Customer, Pitch, Contacts, Quantity, PartNumber, Status, ShipDate) VALUES ('" & txt_name & "', '" & txt_pitch & "', '" & txt_contacts & "', '" & txt_quantity & "', '" & txt_partnumber & "', '" & txt_status & "', '" & txt_shipby & "')"
 MsgBox txt_name & " added to orders"


End Sub

Private Sub btnHome_Click()

DoCmd.Close
DoCmd.OpenForm "Home Page"
End Sub

Private Sub Form_Load()

[Forms]![Add_Element_Order_Form]![statusBox] = "Open"
[Forms]![Add_Element_Order_Form]![pitchBox] = "0.5"
[Forms]![Add_Element_Order_Form]![numberBox] = "ZFLAT-CTF07-"

End Sub
