# ExcelValidator

## Objetivo

O **ExcelValidator** e um serviço desenvolvido em .NET 9 que permite validar arquivos Excel (`.xlsx`) com base em um conjunto flexível de regras. O serviço processa o arquivo Excel, aplica as regras de validaçao definidas para cada coluna e retorna o arquivo modificado, destacando as celulas invalidas com fundo rosa claro e comentarios explicativos.

## Como funciona

  - O serviço recebe um arquivo Excel codificado em Base64 e uma lista de regras de validaçao.
  - Cada regra e associada a uma coluna específica do Excel.
  - Para cada celula, todas as regras definidas para a respectiva coluna sao aplicadas.
  - Se uma celula nao atender a uma ou mais regras, ela e destacada e recebe um comentario com as mensagens de erro.
  - O arquivo Excel validado e retornado tambem em Base64.

## Exemplo de uso

### Estrutura da requisiçao

A requisiçao deve conter:

  - `Base64Excel`: string com o arquivo Excel codificado em Base64.
  - `Rules`: lista de regras de validaçao, onde cada regra possui:
  - `ColumnName`: nome da coluna a ser validada (deve coincidir com o cabeçalho do Excel).
  - `RuleType`: tipo da regra (ex: `Required`, `Regex`, `MaxLength`, `MinLength`, `Numeric`, etc).
  - `RuleValue`: valor de referencia para a regra (quando aplicavel).
  - `ErrorMessage`: mensagem de erro personalizada.

    #### Exemplo de definiçao de regras
   
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
            ErrorMessage = "O e-mail e obrigatorio."
        },
        new ValidationRule
        {
            ColumnName = "Email",
            RuleType = "Regex",
            RuleValue = @"^[\w.-]+@[\w.-]+\.\w+$",
            ErrorMessage = "Formato de e-mail invalido."
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
            ErrorMessage = "O nome deve ter no maximo 50 caracteres."
        }
    }
};
```
#### Observaçao:  
> Voce pode definir múltiplas regras para a mesma coluna, e todas serao aplicadas. Caso uma celula viole mais de uma regra, todas as mensagens de erro serao exibidas no comentario da celula.

### Tipos de regras suportadas

- `Required`: valor obrigatorio (nao pode ser vazio).
- `Regex`: validaçao por expressao regular.
- `MaxLength`: tamanho maximo do texto.
- `MinLength`: tamanho mínimo do texto.
- `Length`: tamanho exato do texto.
- `GTIN`: validaçao de codigo GTIN/EAN.
- `Numeric`: valor numerico.
- `NumericGTZero`: valor numerico maior que zero.

## Como executar

1. Adicione a referencia ao pacote [ClosedXML](https://github.com/ClosedXML/ClosedXML).
2. Importe e utilize o serviço `ExcelValidationService` em seu projeto.
3. Envie a requisiçao conforme o exemplo acima.
4. O resultado sera uma string Base64 do arquivo Excel validado, pronto para download ou exibiçao.
