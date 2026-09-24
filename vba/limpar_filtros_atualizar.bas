Attribute VB_Name = "modDashboardRH"
Option Explicit

' Limpa filtros das tabelas da pasta e atualiza conexões, tabelas dinâmicas e fórmulas.
' Associe esta macro a um botão na aba Dashboard.
Public Sub LimparFiltrosEAtualizar()
    Dim ws As Worksheet
    Dim tabela As ListObject
    Dim tabelaDinamica As PivotTable
    Dim conexao As WorkbookConnection

    On Error Resume Next

    For Each ws In ThisWorkbook.Worksheets
        For Each tabela In ws.ListObjects
            If tabela.AutoFilter.FilterMode Then
                tabela.AutoFilter.ShowAllData
            End If
        Next tabela

        For Each tabelaDinamica In ws.PivotTables
            tabelaDinamica.ClearAllFilters
            tabelaDinamica.RefreshTable
        Next tabelaDinamica
    Next ws

    For Each conexao In ThisWorkbook.Connections
        conexao.Refresh
    Next conexao

    ThisWorkbook.RefreshAll
    Application.CalculateFull

    On Error GoTo 0
    MsgBox "Filtros limpos e relatório atualizado.", vbInformation, "Dashboard RH"
End Sub
