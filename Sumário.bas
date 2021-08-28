Attribute VB_Name = "Sumário"
Sub Sumário()

Dim Plan_TT, Plan_Atual As Integer

'Sheets.Add.Name = "Sumario"

Plan_TT = Sheets.Count
Plan_Atual = 1


Do While Plan_Atual <= Plan_TT


Cells(Plan_Atual, 1).Activate
ActiveCell.Value = Sheets(Plan_Atual).Name

Link_Ativo = "'" & Sheets(Plan_Atual).Name & "'!A1"
Nome_Link_Ativo = Sheets(Plan_Atual).Name


ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=Link_Ativo, TextToDisplay:=Nome_Link_Ativo

Plan_Atual = Plan_Atual + 1


Loop

End Sub

