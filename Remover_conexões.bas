Attribute VB_Name = "Remover_conex�es"
Sub RemoverConex�es()
Dim connection As WorkbookConnection
On Error Resume Next
For Each connection In ThisWorkbook.Connections
connection.Delete
Next
End Sub
