Sub Inserir_RetornoDeObra_Balanco()
    Set wsRetornoDeObra = ThisWorkbook.Sheets("RetornoDeObra")
    Set wsRegEntrada = ThisWorkbook.Sheets("RegEntrada")
    Set wsBalanco = ThisWorkbook.Sheets("Balanço")
    Set tbBalanco = wsBalanco.ListObjects("Balanço")
    Set rngRegistros = wsRetornoDeObra.Range("G3:H" & wsRetornoDeObra.Cells(wsRetornoDeObra.Rows.Count, "G").End(xlUp).Row)

    ' | Parte 1: Transferencia dos IDs dos registros |
    startRow = (wsRegEntrada.Cells(wsRegEntrada.Rows.Count, "A").End(xlUp).Offset(1, 0).Row) - (rngRegistros.Rows.Count)
    endRow = wsRegEntrada.Cells(wsRegEntrada.Rows.Count, "A").End(xlUp).Row
    Set rngATrans_id = wsRegEntrada.Range("A" & startRow & ":A" & endRow)

    Dim clmnId_Operacao As ListColumn
    Set clmnId_Operacao = tbBalanco.ListColumns("Id_Operacao")
    
    If Not clmnId_Operacao.DataBodyRange Is Nothing Then
        clmnId_Operacao.DataBodyRange.Rows(clmnId_Operacao.DataBodyRange.Rows.Count + 1).Resize(rngATrans_id.Rows.Count).Value = rngATrans_id.Value
    Else
        wsBalanco.Cells(2, 3).Resize(rngATrans_id.Rows.Count, 1).Value = rngATrans_id.Value
        wsBalanco.Cells(2, 3).Resize(rngATrans_id.Rows.Count, 1).Value = rngATrans_id.Value
    End If

    ' | Parte 2: Seta o tipo de operação |
    Dim clmnOperacao As ListColumn
    Set clmnOperacao = tbBalanco.ListColumns("Operacao")
    clmnOperacao.DataBodyRange.Rows((clmnOperacao.DataBodyRange.Rows.Count - (endRow - startRow)) & ":" & clmnOperacao.DataBodyRange.Rows.Count).Value = "Entrada"

    ' | Parte 3: Seta ID |
    Dim clmnID As ListColumn
    Set clmnID = tbBalanco.ListColumns("Id")

    If tbBalanco.ListRows.Count <> 0 Then
        For i = clmnID.DataBodyRange.Rows.Count To 1 Step -1
            If IsEmpty(clmnID.DataBodyRange.Cells(i, 1).Value) Then
                clmnID.DataBodyRange.Cells(i, 1).Value = i
            Else
                Exit For
            End If
        Next i
    Else
        For i = 1 To clmnID.DataBodyRange.Rows.Count
            If IsEmpty(clmnID.DataBodyRange.Cells(i, 1).Value) Then
                clmnID.DataBodyRange.Cells(i, 1).Value = i
            Else
                Exit For
            End If
        Next i
    End If
End Sub
