Attribute VB_Name = "Módulo2"
Sub LimparDados()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Definir a planilha
    Set ws = ThisWorkbook.Sheets("Consulta") ' Nome da sua planilha
    
    ' Encontrar a última linha com dados na coluna A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Verificar se há dados a partir da linha 3
    If lastRow >= 3 Then
        ' Limpar o conteúdo das linhas a partir da linha 3 até a última linha
        ws.Rows("3:" & lastRow).ClearContents
    End If
End Sub

