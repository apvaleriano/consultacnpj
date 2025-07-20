Attribute VB_Name = "M�dulo2"
Sub LimparDados()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Definir a planilha
    Set ws = ThisWorkbook.Sheets("Consulta") ' Nome da sua planilha
    
    ' Encontrar a �ltima linha com dados na coluna A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Verificar se h� dados a partir da linha 3
    If lastRow >= 3 Then
        ' Limpar o conte�do das linhas a partir da linha 3 at� a �ltima linha
        ws.Rows("3:" & lastRow).ClearContents
    End If
End Sub

