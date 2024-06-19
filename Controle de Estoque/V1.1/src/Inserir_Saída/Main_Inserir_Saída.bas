Sub Main_Inserir_Saída()
    Set wsSaída = ThisWorkbook.Sheets("Saída")
    Set wsRegSaída = ThisWorkbook.Sheets("RegSaída")
    Set wsBalanco = ThisWorkbook.Sheets("Balanço")
    Set tbBalanco = wsBalanco.ListObjects("Balanço")
    Set tbRegSaída = wsRegSaída.ListObjects("RegSaída")
    Set tbSaída = wsSaída.ListObjects("Saída")
    Set rngRegistros = wsSaída.Range("E3:F" & wsSaída.Cells(wsSaída.Rows.Count, "E").End(xlUp).Row)

    ' Verifica se a célula VERIFICADORA na planilha contém o valor "OK!"
    If wsSaída.Range("C9").Value <> "OK!" Then
        MsgBox "Erro: Favor, verificar 'STATUS'!", vbExclamation
        Range("C2").Select
        Exit Sub
    End If

    ' Chama as sub-rotinas para transferir dados
    Call Inserir_Saída_RegSaída
    Call Inserir_Saída_Balanco

    ' Inserção do DateTime nas Planilhas
    strRngSetDateTimeRegSaída = ((tbRegSaída.ListRows.Count) - (rngRegistros.Rows.Count) + 1) & ":" & (tbRegSaída.ListRows.Count)
    
    tbRegSaída.DataBodyRange.Columns(2).Rows(strRngSetDateTimeRegSaída).Value = Now

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
    tbSaída.ListColumns(1).DataBodyRange.ClearContents
    tbSaída.ListColumns(2).DataBodyRange.ClearContents
    tbSaída.ListColumns(4).DataBodyRange.ClearContents

    wsSaída.Range("C2:C7").ClearContents

    For i = tbSaída.ListRows.Count To 2 Step -1
        tbSaída.ListRows(i).Delete
    Next i
End Sub
