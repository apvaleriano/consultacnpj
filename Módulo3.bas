Attribute VB_Name = "M�dulo3"
Sub ExportarColunasSelecionadas()
    Dim ws As Worksheet
    Dim savePath As String
    Dim lastRow As Long
    Dim wbTemp As Workbook
    Dim wsTemp As Worksheet
    Dim rng As Range
    Dim fileName As String
    Dim resposta As String
    Dim colunasSelecionadas As String
    Dim colArray() As String
    Dim i As Integer, colNum As Integer
    Dim exportRange As Range
    Dim colStart As Long
    Dim colEnd As Long
    
    ' Definir a planilha e encontrar a �ltima linha com dados
    Set ws = ThisWorkbook.Sheets("Consulta") ' Altere "Consulta" pelo nome da sua planilha
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Verificar se h� dados preenchidos na planilha
    If lastRow < 3 Then
        MsgBox "N�o h� dados preenchidos para exportar.", vbExclamation
        Exit Sub
    End If
    
    ' Perguntar ao usu�rio quais colunas ele deseja exportar (separadas por v�rgula)
    colunasSelecionadas = InputBox("Digite os n�meros das colunas que deseja exportar, separados por v�rgula." & vbCrLf & _
                                   "Exemplo: 1, 2, 5, 6", "Sele��o de Colunas")
    
    ' Verificar se o usu�rio inseriu algo
    If colunasSelecionadas = "" Then
        MsgBox "Nenhuma coluna selecionada!", vbCritical
        Exit Sub
    End If
    
    ' Dividir a string das colunas em um array
    colArray = Split(colunasSelecionadas, ",")
    
    ' Inicializar o intervalo de exporta��o
    Set exportRange = Nothing
    For i = LBound(colArray) To UBound(colArray)
        colNum = Val(Trim(colArray(i)))
        If colNum > 0 Then
            If exportRange Is Nothing Then
                Set exportRange = ws.Columns(colNum)
            Else
                Set exportRange = Union(exportRange, ws.Columns(colNum))
            End If
        End If
    Next i
    
    ' Verificar se o intervalo de exporta��o foi definido
    If exportRange Is Nothing Then
        MsgBox "Nenhuma coluna v�lida foi selecionada.", vbCritical
        Exit Sub
    End If
    
    ' Perguntar qual formato a pessoa deseja exportar
    resposta = InputBox("Escolha o formato de exporta��o: " & vbCrLf & "1 - Excel (.xlsx)" & vbCrLf & "2 - CSV (.csv)" & vbCrLf & "3 - PDF (.pdf)", "Exportar Dados")
    
    ' Caminho para salvar os arquivos
    savePath = Application.ThisWorkbook.Path & "\" ' Salvando na mesma pasta do arquivo atual
    
    ' Nome do arquivo (sem extens�o)
    fileName = "Exportacao_" & Format(Now, "yyyymmdd_HHMMSS") ' Exemplo: Exportacao_20240905_120000
    
    ' 1. Se escolher Excel (.xlsx)
    If resposta = "1" Then
        Set wbTemp = Workbooks.Add
        Set wsTemp = wbTemp.Sheets(1)
        exportRange.Copy
        wsTemp.Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        wbTemp.SaveAs savePath & fileName & ".xlsx", FileFormat:=51 ' .xlsx
        wbTemp.Close False
        MsgBox "Arquivo Excel exportado com sucesso!"
    
    ' 2. Se escolher CSV (.csv)
    ElseIf resposta = "2" Then
        Set wbTemp = Workbooks.Add
        Set wsTemp = wbTemp.Sheets(1)
        exportRange.Copy
        wsTemp.Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        wbTemp.SaveAs savePath & fileName & ".csv", FileFormat:=6 ' .csv
        wbTemp.Close False
        MsgBox "Arquivo CSV exportado com sucesso!"
    
    ' 3. Se escolher PDF (.pdf)
    ElseIf resposta = "3" Then
        ' Criar uma nova planilha tempor�ria para exportar como PDF
        Set wbTemp = Workbooks.Add
        Set wsTemp = wbTemp.Sheets(1)
        exportRange.Copy
        wsTemp.Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        
        ' Ajustar a largura das colunas para caber o conte�do
        wsTemp.Columns.AutoFit
        
        ' Definir as configura��es da p�gina para caber tudo em uma p�gina horizontal
        With wsTemp.PageSetup
            .Orientation = xlLandscape ' Modo paisagem
            .Zoom = False ' Desativar o zoom para ajustar
            .FitToPagesWide = 1 ' Ajustar para caber em uma p�gina de largura
            .FitToPagesTall = False ' Altura pode ser indefinida
            .PaperSize = xlPaperA4 ' Tamanho do papel A4
        End With
        
        ' Exportar para PDF
        wsTemp.ExportAsFixedFormat Type:=xlTypePDF, fileName:=savePath & fileName & ".pdf", Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        
        wbTemp.Close False
        MsgBox "Arquivo PDF exportado com sucesso!"
        
    ' Caso o usu�rio insira uma resposta inv�lida
    Else
        MsgBox "Op��o inv�lida! Por favor, escolha 1, 2 ou 3.", vbCritical
    End If
End Sub


