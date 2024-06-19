Sub Main_Inserir_Conferencia()
    Set wsConferencia = ThisWorkbook.Sheets("Conferência")
    Set wsRegMateriaisEntregues = ThisWorkbook.Sheets("RegMateriaisEntregues")
    Set wsRegEntrada = ThisWorkbook.Sheets("RegEntrada")
    Set wsBalanco = ThisWorkbook.Sheets("Balanço")
    Set tbBalanco = wsBalanco.ListObjects("Balanço")
    Set tbRegMateriaisEntregues = wsRegMateriaisEntregues.ListObjects("RegMateriaisEntregues")
    Set tbRegEntrada = wsRegEntrada.ListObjects("RegEntrada")
    Set tbConferencia = wsConferencia.ListObjects("Conferência")
    Set rngRegistros = wsConferencia.Range("E3:J" & wsConferencia.Cells(wsConferencia.Rows.Count, "E").End(xlUp).Row)

    ' Verifica se a célula VERIFICADORA na planilha contém o valor "OK!"
    If wsConferencia.Range("C10").Value <> "OK!" Then
        MsgBox "Erro: Favor, verificar 'STATUS'!", vbExclamation
        Range("C2").Select
        Exit Sub
    End If

    ' Chama as sub-rotinas para transferir dados
    Call Inserir_Conferencia_RegMateriaisEntregues
    Call Inserir_Conferencia_RegEntrada
    Call Inserir_Conferencia_Balanco

    ' Inserção do DateTime nas Planilhas
    strRngSetDateTimeRegMateriaisEntregues = ((tbRegMateriaisEntregues.ListRows.Count) - (rngRegistros.Rows.Count) + 1) & ":" & (tbRegMateriaisEntregues.ListRows.Count)

    tbRegMateriaisEntregues.DataBodyRange.Columns(2).Rows(strRngSetDateTimeRegMateriaisEntregues).Value = Now

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

    wsConferencia.Range("C2:C8").ClearContents

    For i = tbConferencia.ListRows.Count To 2 Step -1
        tbConferencia.ListRows(i).Delete
    Next i
End Sub
