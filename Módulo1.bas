Attribute VB_Name = "Módulo1"
Sub ConsultaCNPJBatchBrasilAPI()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim CNPJ As String
    Dim cleanCNPJ As String
    Dim url As String
    Dim http As Object
    Dim JSON As Object
    Dim processedCount As Long
    
    'Desabilitar para melhorar performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Definir a planilha e encontrar a última linha com dados
    Set ws = ThisWorkbook.Sheets("Consulta")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Loop por cada CNPJ na coluna A
    For i = 3 To lastRow
        ' Atualizar status a cada 10 linhas
        If i Mod 10 = 0 Then
            Application.StatusBar = "Processando CNPJs... " & i & " de " & lastRow
            DoEvents ' Permite que o Excel responda
        End If
        
        CNPJ = ws.Cells(i, 1).Text
        
        ' Limpar CNPJ
        cleanCNPJ = Replace(Replace(Replace(Replace(Replace(CNPJ, "'", ""), ".", ""), "-", ""), "/", ""), " ", "")
        
        If Len(cleanCNPJ) = 14 Then
            url = "https://brasilapi.com.br/api/cnpj/v1/" & cleanCNPJ
            
            Set http = CreateObject("MSXML2.XMLHTTP")
            http.Open "GET", url, False
            http.send
            
            If http.status = 200 Then
                On Error Resume Next
                Set JSON = JsonConverter.ParseJson(http.responseText)
                
                If Not JSON.Exists("message") Then
                    With ws
                        .Cells(i, 2).Value = IIf(IsNull(JSON("razao_social")), "N/A", JSON("razao_social"))
                        .Cells(i, 3).Value = IIf(IsNull(JSON("nome_fantasia")), "N/A", JSON("nome_fantasia"))
                        .Cells(i, 4).Value = IIf(IsNull(JSON("descricao_tipo_logradouro")), "N/A", JSON("descricao_tipo_logradouro")) & " " & IIf(IsNull(JSON("descricao_tipo_de_logradouro")), "N/A", JSON("descricao_tipo_de_logradouro"))
                        .Cells(i, 5).Value = IIf(IsNull(JSON("descricao_tipo_logradouro")), "N/A", JSON("descricao_tipo_logradouro")) & " " & IIf(IsNull(JSON("logradouro")), "N/A", JSON("logradouro"))
                        .Cells(i, 6).Value = IIf(IsNull(JSON("numero")), "N/A", JSON("numero"))
                        .Cells(i, 7).Value = IIf(IsNull(JSON("bairro")), "N/A", JSON("bairro"))
                        .Cells(i, 8).Value = IIf(IsNull(JSON("municipio")), "N/A", JSON("municipio"))
                        .Cells(i, 9).Value = IIf(IsNull(JSON("uf")), "N/A", JSON("uf"))
                        .Cells(i, 10).Value = IIf(IsNull(JSON("cep")), "N/A", JSON("cep"))
                        .Cells(i, 11).Value = IIf(IsNull(JSON("ddd_telefone_1")), "N/A", JSON("ddd_telefone_1"))
                        .Cells(i, 12).Value = IIf(IsNull(JSON("descricao_situacao_cadastral")), "N/A", JSON("descricao_situacao_cadastral"))
                        
                        ' Formatação do Simples Nacional
                        Dim statusSimples As String
                        statusSimples = IIf(IsNull(JSON("opcao_pelo_simples")), "N/A", JSON("opcao_pelo_simples"))
                        
                        If UCase(statusSimples) = "VERDADEIRO" Or statusSimples = "true" Then
                            .Cells(i, 13).Value = "Optante por Simples"
                        Else
                            .Cells(i, 13).Value = "Não Optante"
                        End If
                        
                        .Cells(i, 14).Value = IIf(IsNull(JSON("data_opcao_pelo_simples")), "N/A", JSON("data_opcao_pelo_simples"))
                        .Cells(i, 15).Value = IIf(IsNull(JSON("data_exclusao_do_simples")), "N/A", JSON("data_exclusao_do_simples"))
                        
                        ' Formatação do MEI
                        Dim statusMEI As String
                        statusMEI = IIf(IsNull(JSON("opcao_pelo_mei")), "N/A", JSON("opcao_pelo_mei"))
                        
                        If UCase(statusMEI) = "VERDADEIRO" Or statusMEI = "true" Then
                            .Cells(i, 16).Value = "Optante por MEI"
                        Else
                            .Cells(i, 16).Value = "Não Optante"
                        End If
                        
                        .Cells(i, 17).Value = IIf(IsNull(JSON("data_opcao_pelo_mei")), "N/A", JSON("data_opcao_pelo_mei"))
                        .Cells(i, 18).Value = IIf(IsNull(JSON("data_exclusao_do_mei")), "N/A", JSON("data_exclusao_do_mei"))
                        .Cells(i, 19).Value = IIf(IsNull(JSON("natureza_juridica")), "N/A", JSON("natureza_juridica"))
                        .Cells(i, 20).Value = IIf(IsNull(JSON("codigo_municipio")), "N/A", JSON("codigo_municipio"))
                        .Cells(i, 21).Value = IIf(IsNull(JSON("codigo_municipio_ibge")), "N/A", JSON("codigo_municipio_ibge"))
                        .Cells(i, 22).Value = IIf(IsNull(JSON("cnae_fiscal")), "N/A", JSON("cnae_fiscal"))
                        .Cells(i, 23).Value = IIf(IsNull(JSON("descricao_identificador_matriz_filial")), "N/A", JSON("descricao_identificador_matriz_filial"))
                        .Cells(i, 24).Value = IIf(IsNull(JSON("porte")), "N/A", JSON("porte"))
                    End With
                    
                    ' CNAEs Secundários
                    If Not IsNull(JSON("cnaes_secundarios")) Then
                        Dim cnaesArray As Object
                        Dim cnaeCodigo As String
                        Dim j As Long
                        
                        Set cnaesArray = JSON("cnaes_secundarios")
                        cnaeCodigo = ""
                        For j = 1 To cnaesArray.Count
                            If Len(cnaeCodigo) > 0 Then
                                cnaeCodigo = cnaeCodigo & ", "
                            End If
                            cnaeCodigo = cnaeCodigo & cnaesArray(j)("codigo")
                        Next j
                        ws.Cells(i, 25).Value = cnaeCodigo
                    Else
                        ws.Cells(i, 25).Value = "N/A"
                    End If
                                        
                Else
                    ws.Cells(i, 2).Value = "Erro na consulta"
                    ws.Cells(i, 3).Value = JSON("message")
                End If
                
            ElseIf http.status = 429 Then
                ws.Cells(i, 2).Value = "Erro na consulta"
                ws.Cells(i, 3).Value = "Muitas requisições, tente novamente mais tarde"
                
                ' Esperar 5 segundos se receber erro de muitas requisições
                Application.Wait Now + TimeSerial(0, 0, 5)
                
            Else
                ws.Cells(i, 2).Value = "Erro na consulta"
                ws.Cells(i, 3).Value = "Verifique o CNPJ"
            End If
            
            ' Pequena pausa entre requisições para evitar sobrecarga
            Application.Wait Now + TimeSerial(0, 0, 1)
            
        Else
            ws.Cells(i, 2).Value = "CNPJ Inválido"
        End If
    Next i
    
    ' Restaurar configurações
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Processamento concluído!", vbInformation
End Sub
