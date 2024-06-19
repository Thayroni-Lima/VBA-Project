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
    strRngSetDadosUnicos = ((tbRegSaída.ListRows.Count) - (rngUnionRegistros.Rows.Count) + 1) & ":" & (tbRegSaída.ListRows.Count)

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
