Attribute VB_Name = "Remover_conexões"
Sub RemoverConexões()
Dim connection As WorkbookConnection
On Error Resume Next
For Each connection In ThisWorkbook.Connections
connection.Delete
Next
End Sub
