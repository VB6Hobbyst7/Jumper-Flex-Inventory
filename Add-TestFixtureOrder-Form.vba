Option Compare Database

Private Sub addOrder_Click()

End Sub

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
Dim txt_flexlength As String
Dim txt_status As String
Dim txt_shipby As String
Dim flexQuery As DAO.Recordset
Dim flexes_in_stock As Variant
Dim drawing_answer As Integer
Dim short As Integer

' feilds for sending email
Dim OutApp As Object
Dim OutMail As Object
Dim strbody As String
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

 
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

' generate string to send in email if the flex DNE
strbody = "Hi John" & vbNewLine & vbNewLine & _
          "Please order this flex:" & vbNewLine & _
          "Pitch: " & txt_pitch & vbNewLine & _
          "Contacts: " & txt_contacts & vbNewLine & _
          "Length: " & txt_flexlength & " inches " & vbNewLine & _
          "Quantity: " & txt_quantity & vbNewLine & _
          "Customer: " & txt_name

' Running a SQL to check for flexes in stock -- for now
' flexSQL = "SELECT [Copy Of Jumper Flex Inventory--PRODUCTION].Pitch_mm, [Copy Of Jumper Flex Inventory--PRODUCTION].[Number of Conductors], [Copy Of Jumper Flex Inventory--PRODUCTION].Length_in FROM [Copy Of Jumper Flex Inventory--PRODUCTION] WHERE ((([Copy Of Jumper Flex Inventory--PRODUCTION].Pitch_mm)=  '" & txt_pitch & "' And (([Copy Of Jumper Flex Inventory--PRODUCTION].[Number of Conductors])= '" & txt_contacts & "') And (([Copy Of Jumper Flex Inventory--PRODUCTION].Length_in)= '" & txt_flexlength & "'));"

DoCmd.RunSQL "INSERT INTO CTF_Open_Orders (Customer, Pitch, Contacts, Quantity, PartNumber, Status, ShipDate) VALUES ('" & txt_name & "', '" & txt_pitch & "', '" & txt_contacts & "', '" & txt_quantity & "', '" & txt_partnumber & "', '" & txt_status & "', '" & txt_shipby & "')"

answer = MsgBox("Did we get a drawing?", vbQuestion + vbYesNo + vbDefaultButton2)
short = MsgBox("Is the flex short access?", vbQuestion + vbYesNo + vbDefaultButton2)

Form.Refresh

' count if there are any flexes returned by the query
' somewhat works but if there is nothing there -- it'll just return null
flexes_in_stock = DSum("Quantity", "checkFlexesAfterOrder_Query")

' if the the quantity value is 0, email to JG the correct flexes to use
If IsNull(flexes_in_stock) Then
    ' MsgBox "Please Order " & txt_pitch & "mm x " & txt_contacts & " x " & txt_flexlength & " inch flex." & vbNewLine & "Quantity " & txt_quantity
    On Error Resume Next
    With OutMail
        .To = "johnglatts1@hotmail.com"
        .CC = ""
        .BCC = ""
        .Subject = "This is the Subject line"
        .Body = strbody
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
        .Send   'or use .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
Else
    lblCheckFlex.Visible = True
    Label97.Visible = True
    checkFlexesAfterOrder_QuerySubform.Visible = True
End If

End Sub

Private Sub btnHome_Click()

DoCmd.Close
DoCmd.OpenForm "Home Page"

End Sub

Private Sub Command94_Click()
' Testing looping through the flex table


End Sub

Private Sub Form_Load()

statusBox.Value = "Open"

lblCheckFlex.Visible = False
Label97.Visible = False
checkFlexesAfterOrder_QuerySubform.Visible = False



End Sub
