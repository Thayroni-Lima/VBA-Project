' Declaração de variáveis globais
Dim wsSaída As Worksheet
Dim wsRegSaída As Worksheet
Dim wsBalanco As Worksheet
Dim tbSaída As ListObject
Dim tbRegSaída As ListObject
Dim tbBalanco As ListObject
Sub Inserir_Saída_Balanco()
    Set wsSaída = ThisWorkbook.Sheets("Saída")
    Set wsRegSaída = ThisWorkbook.Sheets("RegSaída")
    Set wsBalanco = ThisWorkbook.Sheets("Balanço")
    Set tbBalanco = wsBalanco.ListObjects("Balanço")
    Set rngRegistros = wsSaída.Range("E3:H" & wsSaída.Cells(wsSaída.Rows.Count, "E").End(xlUp).Row)

    ' | Parte 1: Transferencia dos IDs dos registros |
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
Sub Inserir_Saída_RegSaída()
    ' | Parte 1: Transferência de Registros |
    Set wsSaída = ThisWorkbook.Sheets("Saída")
    Set wsRegSaída = ThisWorkbook.Sheets("RegSaída")
    Set tbRegSaída = wsRegSaída.ListObjects("RegSaída")
    Set rngRegistros = wsSaída.Range("E3:F" & wsSaída.Cells(wsSaída.Rows.Count, "E").End(xlUp).Row)

    Set rngRegistrosObs = wsSaída.Range("H3:H" & wsSaída.Cells(wsSaída.Rows.Count, "E").End(xlUp).Row)
    Set rngUnionRegistros = Union(rngRegistros, rngRegistrosObs)

    rngUnionRegistros.Copy

    Dim clmnMaterial_Retirado As ListColumn
    Set clmnMaterial_Retirado = tbRegSaída.ListColumns("Material_Retirado")

    If Not clmnMaterial_Retirado.DataBodyRange Is Nothing Then
        wsRegSaída.Cells(wsRegSaída.Rows.Count, "I").End(xlUp).Offset(1, 0).Resize(rngUnionRegistros.Rows.Count, rngUnionRegistros.Columns.Count).PasteSpecial Paste:=xlPasteValues
    Else
        wsRegSaída.Cells(2, "I").Resize(rngUnionRegistros.Rows.Count, rngUnionRegistros.Columns.Count).PasteSpecial Paste:=xlPasteValues
    End If

    ' | Parte 2: Transferência de dados únicos (data, hora, etc.) |
    Set trRngSetDadosUnicos = ((tbRegSaída.ListRows.Count) - (rngRegistros.Rows.Count) + 1) & ":" & (tbRegSaída.ListRows.Count)

    For i = 3 To 8
        tbRegSaída.DataBodyRange.Columns(i).Rows(strRngSetDadosUnicos).Value = wsSaída.Range("C" & (i - 1)).Value
    Next i

    ' | Parte 3: Seta ID |
    Dim clmnID As ListColumn
    Set clmnID = tbRegSaída.ListColumns("Id")

    If tbRegSaída.ListRows.Count <> 0 Then
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
    
    ' Desativa o modo de corte e copia do Excel
    Application.CutCopyMode = False
End Sub
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
