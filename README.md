
# Consulta CNPJ

Consulta CNPJ Excel VBA
Uma macro em VBA para Excel que automatiza a consulta em lote de CNPJs.


📋 Funcionalidades

Consulta em lote de CNPJs
Extração automática de dados como:
Razão Social
Nome Fantasia
Endereço completo
Situação cadastral
Informações sobre Simples Nacional e MEI
CNAEs (principal e secundários)
Dados municipais e outras informações cadastrais



🚀 Como usar

Abra sua planilha do Excel
Importe o módulo VBA com o código
Prepare uma planilha chamada "Consulta" com os CNPJs na coluna A
Execute a macro ConsultaCNPJBatchBrasilAPI()
## Screenshots

![App Screenshot](https://i.imgur.com/fp9TEnT.png)


## 

🚦 Controle de Requisições

Intervalo de 1 segundo entre requisições para evitar sobrecarga
Tratamento de erro 429 (muitas requisições) com pausa de 5 segundos


🛠️ Tratamentos Especiais

Limpeza automática de formatação do CNPJ
Validação do tamanho do CNPJ (14 dígitos)
Tratamento de valores nulos/ausentes


⚠️ Limitações

Sujeito aos limites de requisição da API
Necessita de conexão com a internet

📝 Notas

Recomenda-se verificar os termos de uso da API antes de utilizar em ambiente de produção
Para grandes volumes de consulta, considere implementar um sistema de cache

🤝 Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para:
Reportar bugs
Sugerir novas funcionalidades
Melhorar a documentação
