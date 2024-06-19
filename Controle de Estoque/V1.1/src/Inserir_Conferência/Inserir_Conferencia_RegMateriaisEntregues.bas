Sub Inserir_Conferencia_RegMateriaisEntregues()
    ' | Parte 1: Transferência dos Regitros |
    Set wsConferencia = ThisWorkbook.Sheets("Conferência")
    Set wsRegMateriaisEntregues = ThisWorkbook.Sheets("RegMateriaisEntregues")
    Set tbRegMateriaisEntregues = wsRegMateriaisEntregues.ListObjects("RegMateriaisEntregues")
    Set rngRegistros = wsConferencia.Range("G3:J" & wsConferencia.Cells(wsConferencia.Rows.Count, "G").End(xlUp).Row)

    rngRegistros.Copy

    Dim clmnMaterial_Entregue As ListColumn
    Set clmnMaterial_Entregue = tbRegMateriaisEntregues.ListColumns("Material_Entregue")
    
    If Not clmnMaterial_Entregue.DataBodyRange Is Nothing Then
        wsRegMateriaisEntregues.Cells(wsRegMateriaisEntregues.Rows.Count, "J").End(xlUp).Offset(1, 0).Resize(rngRegistros.Rows.Count, rngRegistros.Columns.Count).PasteSpecial Paste:=xlPasteValues
    Else
        wsRegMateriaisEntregues.Cells(2, "J").Resize(rngRegistros.Rows.Count, rngRegistros.Columns.Count).PasteSpecial Paste:=xlPasteValues
    End If

    ' | Parte 2: Transferência de dados únicos (data, hora, etc.) |
    strRngSetDadosUnicos = ((tbRegMateriaisEntregues.ListRows.Count) - (rngRegistros.Rows.Count) + 1) & ":" & (tbRegMateriaisEntregues.ListRows.Count)

    For i = 3 To 9
        tbRegMateriaisEntregues.DataBodyRange.Columns(i).Rows(strRngSetDadosUnicos).Value = wsConferencia.Range("C" & (i - 1)).Value
    Next i

    ' | Parte 3: Seta ID |
    Dim clmnID As ListColumn
    Set clmnID = tbRegMateriaisEntregues.ListColumns("Id")

    If tbRegMateriaisEntregues.ListRows.Count <> 0 Then
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
    
    ' Desativa o modo de corte e cópia do Excel
    Application.CutCopyMode = False
End Sub
