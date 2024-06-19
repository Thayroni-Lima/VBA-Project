Sub Inserir_Saída_Balanco()
    Set wsSaída = ThisWorkbook.Sheets("Saída")
    Set wsRegSaída = ThisWorkbook.Sheets("RegSaída")
    Set wsBalanco = ThisWorkbook.Sheets("Balanço")
    Set tbBalanco = wsBalanco.ListObjects("Balanço")
    Set rngRegistros = wsSaída.Range("E3:H" & wsSaída.Cells(wsSaída.Rows.Count, "E").End(xlUp).Row)

    ' | Parte 1: Transferencia dos IDs dos registros |
    startRow = (wsRegSaída.Cells(wsRegSaída.Rows.Count, "A").End(xlUp).Offset(1, 0).Row) - (rngRegistros.Rows.Count)
    endRow = wsRegSaída.Cells(wsRegSaída.Rows.Count, "A").End(xlUp).Row
    Set rngATrans_id = wsRegSaída.Range("A" & startRow & ":A" & endRow)

    Dim clmnId_Operacao As ListColumn
    Set clmnId_Operacao = tbBalanco.ListColumns("Id_Operacao")
    
    clmnId_Operacao.DataBodyRange.Rows(clmnId_Operacao.DataBodyRange.Rows.Count + 1).Resize(rngATrans_id.Rows.Count).Value = rngATrans_id.Value

    ' | Parte 2: Seta o tipo de operação |
    Dim clmnOperacao As ListColumn
    Set clmnOperacao = tbBalanco.ListColumns("Operacao")

    strRngSetDadosUnicos = (clmnOperacao.DataBodyRange.Rows.Count - (endRow - startRow)) & ":" & clmnOperacao.DataBodyRange.Rows.Count
    clmnOperacao.DataBodyRange.Rows(strRngSetDadosUnicos).Value = "Saída"

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
