Private Sub Form_Load()
' Update everything when the form is loaded
' Update the material textboxes and the open orders
Dim total_blanks As Integer
Dim orders_1 As String
Dim orders_2 As String
Dim orders_3 As String
Dim orders_4 As String
Dim orders_5 As String
Dim orders_6 As String
Dim total_C_blanks As Integer
Dim total_G4_skins As Integer
Dim total_C_skins As Integer
Dim total_diving_boards As Integer
Dim SQL As String
Dim dbs As DAO.Database
Dim rsStock As DAO.Recordset
Dim Cstock As DAO.Recordset
Dim inventorystock As DAO.Recordset
Dim openOrdersTable As DAO.Recordset

Set dbs = CurrentDb
Set rsStock = dbs.OpenRecordset("G4_Stock", dbOpenSnapshot)
Set Cstock = dbs.OpenRecordset("C_Stock", dbOpenSnapshot)
Set inventorystock = dbs.OpenRecordset("Material_Inventory", dbOpenSnapshot)



' Run the INSERT query to find all open orders
SQL = "SELECT CTF_open_orders.Customer, CTF_open_orders.Status INTO CTF_open_orders_for_display FROM CTF_open_orders WHERE (((CTF_open_orders.Status)='Open'));"
DoCmd.RunSQL SQL

' Wait to open the table until  query is done
Set openOrdersTable = dbs.OpenRecordset("CTF_open_orders_for_display", dbOpenSnapshot)

total_blanks = DSum("Full_Blanks", "G4_Stock")
total_C_blanks = DSum("Full_Blanks", "C_Stock")
total_G4_skins = DSum("G4_Skins", "Material_Inventory")
total_C_skins = DSum("C_Skins", "Material_Inventory")
total_diving_boards = DSum("Diving_Boards", "Material_Inventory")
' testing getting open orders to go in textbox
' somehow get the id's to always be 1-6 in the opeen orders query
' Or use the query to fill up another table, and just the table to get the right ID numbers
orders_1 = DLookup("[Customer]", "CTF_open_orders_for_display", "[ID] = 1")
orders_2 = DLookup("[Customer]", "CTF_open_orders_for_display", "[ID] = 2")
orders_3 = DLookup("[Customer]", "CTF_open_orders_for_display", "[ID] = 3")
orders_4 = DLookup("[Customer]", "CTF_open_orders_for_display", "[ID] = 4")
orders_5 = DLookup("[Customer]", "CTF_open_orders_for_display", "[ID] = 5")
orders_6 = DLookup("[Customer]", "CTF_open_orders_for_display", "[ID] = 6")

txtCTotal = total_C_blanks
txtG4Total = total_blanks
g4_skins_txt = total_G4_skins
c_skins_txt = total_C_skins
diveboard_txt = total_diving_boards
txtOrders1 = orders_1
txtOrders2 = orders_2
txtOrders3 = orders_3
txtOrders4 = orders_4
txtOrders5 = orders_5
txtOrders6 = orders_6
' set focus to button to get rid of random white box
add_inventory_btn.SetFocus


End Sub
