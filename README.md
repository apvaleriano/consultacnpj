
# Consulta CNPJ

Uma macro em VBA para Excel que automatiza a consulta em lote de CNPJs. O sistema formata automaticamente os CNPJs digitados e padroniza as informações de Simples Nacional e MEI.
## 
📋 Funcionalidades

Consulta em lote de CNPJs

Formatação automática de CNPJ na digitação (XX.XXX.XXX/XXXX-XX)

Preenchimento automático com zeros à esquerda

Padronização das informações de Simples Nacional e MEI

Extração automática de dados como:
Razão Social

Nome Fantasia

Endereço completo

Telefone

Situação cadastral

Informações sobre Simples Nacional e MEI

CNAEs (principal e secundários)

Dados municipais e outras informações cadastrais

## 
🚀 Como usar

Abra sua planilha do Excel

Importe o módulo VBA com o código

Prepare uma planilha chamada "Consulta" com os CNPJs na coluna A

Execute a macro ConsultaCNPJBatchBrasilAPI()

## Screenshots

![App Screenshot](https://i.imgur.com/wxK1PpZ.png)


## 

🚦 Controle de Requisições

Intervalo de 1 segundo entre requisições para evitar sobrecarga

Tratamento de erro 429 (muitas requisições) com pausa de 5 segundos
## 

🛠️ Tratamentos Especiais

Limpeza automática de formatação do CNPJ

Validação do tamanho do CNPJ (14 dígitos)

Tratamento de valores nulos/ausentes
## 

📝 Notas

Recomenda-se verificar os termos de uso da API antes de utilizar em ambiente de produção

Para grandes volumes de consulta, considere implementar um sistema de cache
## 
🤝 Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para:

Reportar bugs

Sugerir novas funcionalidades

Melhorar a documentação
