' Declaração de variáveis globais
Dim wsRetornoDeObra As Worksheet
Dim wsRegEntrada As Worksheet
Dim wsBalanco As Worksheet
Dim wsObras As Worksheet
Dim tbRegEntrada As ListObject
Dim tbBalanco As ListObject
Dim tbRetornoDeObra As ListObject
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
Sub Inserir_RetornoDeObra_RegEntrada()
    ' | Parte 1: Transferência de Registros |
    Set wsRetornoDeObra = ThisWorkbook.Sheets("RetornoDeObra")
    Set wsRegEntrada = ThisWorkbook.Sheets("RegEntrada")
    Set wsObras = ThisWorkbook.Sheets("Obras")
    Set tbRegEntrada = wsRegEntrada.ListObjects("RegEntrada")
    Set rngRegistros = wsRetornoDeObra.Range("E3:G" & wsRetornoDeObra.Cells(wsRetornoDeObra.Rows.Count, "E").End(xlUp).Row)

    rngRegistros.Copy

    Dim clmnMaterial_Entregue As ListColumn
    Set clmnMaterial_Entregue = tbRegEntrada.ListColumns("Material_Entregue")

    If Not clmnMaterial_Entregue.DataBodyRange Is Nothing Then
        wsRegEntrada.Cells(wsRegEntrada.Rows.Count, "I").End(xlUp).Offset(1, 0).Resize(rngRegistros.Rows.Count, rngRegistros.Columns.Count).PasteSpecial Paste:=xlPasteValues
    Else
        wsRegEntrada.Cells(2, "I").Resize(rngRegistros.Rows.Count, rngRegistros.Columns.Count).PasteSpecial Paste:=xlPasteValues
    End If

    ' | Parte 2: Transferência de dados únicos (data, hora, etc.) |
    strRngSetDadosUnicos = (((tbRegEntrada.ListRows.Count) - (rngRegistros.Rows.Count)) + 1) & ":" & (tbRegEntrada.ListRows.Count)

    For i = 3 To 8
        If i = 6 Then
            tbRegEntrada.DataBodyRange.Columns(6).Rows(strRngSetDadosUnicos).Value = "Retorno de Obra"
        ElseIf i = 7 Then
            tbRegEntrada.DataBodyRange.Columns(i).Rows(strRngSetDadosUnicos).Value = wsRetornoDeObra.Range("C" & (i - 2)).Value
        ElseIf i = 8 Then
            tbRegEntrada.DataBodyRange.Columns(8).Rows(strRngSetDadosUnicos).Value = wsObras.Cells(2, 2).Value
        Else
            tbRegEntrada.DataBodyRange.Columns(i).Rows(strRngSetDadosUnicos).Value = wsRetornoDeObra.Range("C" & (i - 1)).Value
        End If
    Next i

    ' | Parte 3: Seta ID |
    Dim clmnID As ListColumn
    Set clmnID = tbRegEntrada.ListColumns("Id")

    If tbRegEntrada.ListRows.Count <> 0 Then
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
