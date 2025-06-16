# ExcelValidator

## Objetivo

O **ExcelValidator** é um serviço desenvolvido em .NET 9 que permite validar arquivos Excel (`.xlsx`) com base em um conjunto flexível de regras. O serviço processa o arquivo Excel, aplica as regras de validação definidas para cada coluna e retorna o arquivo modificado, destacando as células inválidas com fundo rosa claro e comentários explicativos.

## Como funciona

- O serviço recebe um arquivo Excel codificado em Base64 e uma lista de regras de validação.
- Cada regra é associada a uma coluna específica do Excel.
- Para cada célula, todas as regras definidas para a respectiva coluna são aplicadas.
- Se uma célula não atender a uma ou mais regras, ela é destacada e recebe um comentário com as mensagens de erro.
- O arquivo Excel validado é retornado também em Base64.

## Exemplo de uso

### Estrutura da requisição

A requisição deve conter:

- `Base64Excel`: string com o arquivo Excel codificado em Base64.
- `Rules`: lista de regras de validação, onde cada regra possui:
  - `ColumnName`: nome da coluna a ser validada (deve coincidir com o cabeçalho do Excel).
  - `RuleType`: tipo da regra (ex: `Required`, `Regex`, `MaxLength`, `MinLength`, `Numeric`, etc).
  - `RuleValue`: valor de referência para a regra (quando aplicável).
  - `ErrorMessage`: mensagem de erro personalizada.

    #### Exemplo de definição de regras
   
 ```csharp
var request = new ValidationRequest
{
    Base64Excel = "<arquivo_excel_em_base64>",
    Rules = new List<ValidationRule>
    {
        new ValidationRule
        {
            ColumnName = "Email",
            RuleType = "Required",
            ErrorMessage = "O e-mail é obrigatório."
        },
        new ValidationRule
        {
            ColumnName = "Email",
            RuleType = "Regex",
            RuleValue = @"^[\w.-]+@[\w.-]+\.\w+$",
            ErrorMessage = "Formato de e-mail inválido."
        },
        new ValidationRule
        {
            ColumnName = "Idade",
            RuleType = "NumericGTZero",
            ErrorMessage = "A idade deve ser um número maior que zero."
        },
        new ValidationRule
        {
            ColumnName = "Nome",
            RuleType = "MaxLength",
            RuleValue = "50",
            ErrorMessage = "O nome deve ter no máximo 50 caracteres."
        }
    }
};
````
#### Observação:  
> Você pode definir múltiplas regras para a mesma coluna, e todas serão aplicadas. Caso uma célula viole mais de uma regra, todas as mensagens de erro serão exibidas no comentário da célula.

### Tipos de regras suportadas

- `Required`: valor obrigatório (não pode ser vazio).
- `Regex`: validação por expressão regular.
- `MaxLength`: tamanho máximo do texto.
- `MinLength`: tamanho mínimo do texto.
- `Length`: tamanho exato do texto.
- `GTIN`: validação de código GTIN/EAN.
- `Numeric`: valor numérico.
- `NumericGTZero`: valor numérico maior que zero.

## Como executar

1. Adicione a referência ao pacote [ClosedXML](https://github.com/ClosedXML/ClosedXML).
2. Importe e utilize o serviço `ExcelValidationService` em seu projeto.
3. Envie a requisição conforme o exemplo acima.
4. O resultado será uma string Base64 do arquivo Excel validado, pronto para download ou exibição.

