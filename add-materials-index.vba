
Public whatsClicked As String
Option Compare Database

Private Sub btnCMoldedSkins_Click()

[lblQuantity].Visible = True
[txtQuantity].Visible = True
[btnUpdate].Visible = True

whatsClicked = "MoldedCSkins"


End Sub


'---------------------------------------------
Private Sub btnCskins_Click()

[lblQuantity].Visible = True
[txtQuantity].Visible = True
[btnUpdate].Visible = True

whatsClicked = "Cskins"


End Sub


'---------------------------------------------
Private Sub btnDivingboards_Click()

[lblQuantity].Visible = True
[txtQuantity].Visible = True
[btnUpdate].Visible = True

whatsClicked = "divingboards"

End Sub


'---------------------------------------------
Private Sub btnG4MoldedSkins_Click()

[lblQuantity].Visible = True
[txtQuantity].Visible = True
[btnUpdate].Visible = True

whatsClicked = "MoldedG4Skins"

End Sub


'---------------------------------------------
Private Sub btnG4skins_Click()

[lblQuantity].Visible = True
[txtQuantity].Visible = True
[btnUpdate].Visible = True

whatsClicked = "G4skins"

End Sub


'---------------------------------------------
Private Sub btnUpdate_Click()
Dim SQL As String
Dim Quantity As Integer
Dim dbs As DAO.Database
Dim rsStock As DAO.Recordset
Dim elementStock As DAO.Recordset

' Open up the Materials Table
Set dbs = CurrentDb
Set rsStock = dbs.OpenRecordset("Material_Inventory", dbOpenSnapshot)
Set rsStock = dbs.OpenRecordset("Material_Inventory", dbOpenSnapshot)
Quantity = [Forms]![add_material_form]![txtQuantity]

    ' Find out what button is selected
    Select Case whatsClicked
    Case "G4skins":
            [txtQuantity].Visible = False
            SQL = "UPDATE [Material_Inventory] SET [Material_Inventory].G4_Skins = [Material_Inventory]![G4_Skins] + " & Quantity & ";"
            DoCmd.RunSQL SQL
            Form.Refresh
            MsgBox "Added: " & Quantity & " new G4 Skins "
            
    Case "MoldedG4Skins":
            ' Add a diff. textbox to prompt if any skins are being removed
            [txtQuantity].Visible = False
            SQL = "UPDATE [G4_Stock] SET [G4_Stock].Full_Blanks = [G4_Stock]![Full_Blanks] + " & Quantity & ";"
            DoCmd.RunSQL SQL
            Form.Refresh
            MsgBox "Added: " & Quantity & " new G4 Skins "
    Case "G4skins":
            [txtQuantity].Visible = False
            SQL = "UPDATE [Material_Inventory] SET [Material_Inventory].G4_Skins = [Material_Inventory]![G4_Skins] + " & Quantity & ";"
            DoCmd.RunSQL SQL
            Form.Refresh
            MsgBox "Added: " & Quantity & " new G4 Skins "
    Case "Cskins":
            [txtQuantity].Visible = False
            SQL = "UPDATE [Material_Inventory] SET [Material_Inventory].C_Skins = [Material_Inventory]![C_Skins] + " & Quantity & ";"
            DoCmd.RunSQL SQL
            Form.Refresh
            MsgBox "Added: " & Quantity & " new C Skins "
    Case "MoldedCskins":
            ' Add a diff. textbox to prompt if any skins are being removed
            [txtQuantity].Visible = False
            SQL = "UPDATE [C_Stock] SET [C_Stock].Full_Blanks = [C_Stock]![Full_Blanks] + " & Quantity & ";"
            DoCmd.RunSQL SQL
            Form.Refresh
            MsgBox "Added: " & Quantity & " new C Skins "
    Case "divingboards":
            [txtQuantity].Visible = False
            SQL = "UPDATE [Material_Inventory] SET [Material_Inventory].Diving_Boards = [Material_Inventory]![Diving_Boards] + " & Quantity & ";"
            DoCmd.RunSQL SQL
            Form.Refresh
            MsgBox "Added: " & Quantity & " new Diving Boards"
    End Select


End Sub


'---------------------------------------------
Private Sub Command29_Click()

DoCmd.Close
DoCmd.OpenForm "Home Page"

End Sub

Private Sub Command30_Click()
Dim MailStr As String
' Not Used

MailStr = ""

If listboxItems.SelectedItems.Count = 0 Then
    MsgBox "No Item Selected"
    Exit Sub
End If
For i = 0 To Me!listboxItems.Items.Count - 1
    If Me!listboxItems.Selected(i) Then
        'MailStr = MailStr & listboxItems.Items.Item(i) & "; "
        MsgBox Me!listboxItems.Items.Item(i)
    End If
Next i

End Sub


'---------------------------------------------
Private Sub Form_Load()

[lblQuantity].Visible = False
[txtQuantity].Visible = False
[btnUpdate].Visible = False

End Sub
