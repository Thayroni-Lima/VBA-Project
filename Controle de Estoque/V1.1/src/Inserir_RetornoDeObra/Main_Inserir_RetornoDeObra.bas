Sub Main_Inserir_RetornoDeObra()
    Set wsRetornoDeObra = ThisWorkbook.Sheets("RetornoDeObra")
    Set wsRegEntrada = ThisWorkbook.Sheets("RegEntrada")
    Set wsBalanco = ThisWorkbook.Sheets("Balanço")
    Set tbRegEntrada = wsRegEntrada.ListObjects("RegEntrada")
    Set tbBalanco = wsBalanco.ListObjects("Balanço")
    Set tbRetornoDeObra = wsRetornoDeObra.ListObjects("RetornoDeObra")
    Set rngRegistros = wsRetornoDeObra.Range("E3:G" & wsRetornoDeObra.Cells(wsRetornoDeObra.Rows.Count, "E").End(xlUp).Row)

    ' Verifica se a célula VERIFICADORA na planilha contém o valor "OK!"
    If wsRetornoDeObra.Range("C7").Value <> "OK!" Then
        MsgBox "Erro: Favor, verificar 'STATUS'!", vbExclamation
        Range("C2").Select
        Exit Sub
    End If

    ' Chama as sub-rotinas para transferir dados
    Call Inserir_RetornoDeObra_RegEntrada
    Call Inserir_RetornoDeObra_Balanco

    ' Inserção do DateTime nas Planilhas
    strRngSetDateTimeRegEntrada = ((tbRegEntrada.ListRows.Count) - (rngRegistros.Rows.Count) + 1) & ":" & (tbRegEntrada.ListRows.Count)
    
    tbRegEntrada.DataBodyRange.Columns(2).Rows(strRngSetDateTimeRegEntrada).Value = Now

    Dim clmnDateTime As ListColumn
    Set clmnDateTime = tbBalanco.ListColumns("DateTime_Registro")
    
    If Application.WorksheetFunction.CountA(clmnDateTime.DataBodyRange) <> 0 Then
        For i = clmnDateTime.DataBodyRange.Rows.Count To 1 Step -1
            If clmnDateTime.DataBodyRange.Rows(i).Value <> "" Then
                clmnDateTime.DataBodyRange.Rows((i + 1) & ":" & clmnDateTime.DataBodyRange.Rows.Count).Value = Now
                Exit For
            End If
        Next i
    Else
        clmnDateTime.DataBodyRange.Rows("1:" & clmnDateTime.DataBodyRange.Rows.Count).Value = Now
    End If

    ' Limpeza do Front
    rngRegistros.ClearContents

    wsRetornoDeObra.Range("C2:C5").ClearContents

    For i = tbRetornoDeObra.ListRows.Count To 2 Step -1
        tbRetornoDeObra.ListRows(i).Delete
    Next i
End Sub
