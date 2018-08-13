Private Sub Command77_Click()
'Update quantity automatically'

Dim SQL As String

SQL = "UPDATE [PCB- Query] SET [PCB- Query].Quantity = [PCB- Query]![Quantity]-[Forms]![FIND-PCB]![Text17];"

DoCmd.RunSQL SQL
DoCmd.OpenForm "APP-START"

End Sub
