
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