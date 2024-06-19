Sub Inserir_Conferencia_RegEntrada()
    ' | Parte 1: Transferência dos Registros |
    Set wsConferencia = ThisWorkbook.Sheets("Conferência")
    Set wsRegEntrada = ThisWorkbook.Sheets("RegEntrada")
    Set tbRegEntrada = wsRegEntrada.ListObjects("RegEntrada")
    Set rngRegistros = wsConferencia.Range("G3:H" & wsConferencia.Cells(wsConferencia.Rows.Count, "G").End(xlUp).Row)

    Set rngRegistrosObs = wsConferencia.Range("J3:J" & wsConferencia.Cells(wsConferencia.Rows.Count, "G").End(xlUp).Row)
    Set rngUnionRegistros = Union(rngRegistros, rngRegistrosObs)

    rngUnionRegistros.Copy

    Dim clmnMaterial_Entregue As ListColumn
    Set clmnMaterial_Entregue = tbRegEntrada.ListColumns("Material_Entregue")

    If Not clmnMaterial_Entregue.DataBodyRange Is Nothing Then
        wsRegEntrada.Cells(wsRegEntrada.Rows.Count, "I").End(xlUp).Offset(1, 0).Resize(rngUnionRegistros.Rows.Count, rngUnionRegistros.Columns.Count).PasteSpecial Paste:=xlPasteValues
    Else
        wsRegEntrada.Cells(2, "I").Resize(rngUnionRegistros.Rows.Count, rngUnionRegistros.Columns.Count).PasteSpecial Paste:=xlPasteValues
    End If

    ' | Parte 2: Transferência de dados únicos (data, hora, etc.) |
    strRngSetDadosUnicos = ((tbRegEntrada.ListRows.Count) - (rngUnionRegistros.Rows.Count) + 1) & ":" & (tbRegEntrada.ListRows.Count)

    For i = 3 To 8
        tbRegEntrada.DataBodyRange.Columns(i).Rows(strRngSetDadosUnicos).Value = wsConferencia.Range("C" & (i - 1)).Value
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
