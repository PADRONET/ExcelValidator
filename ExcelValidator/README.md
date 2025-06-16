# ExcelValidator

## Objetivo

O **ExcelValidator** � um servi�o desenvolvido em .NET 9 que permite validar arquivos Excel (`.xlsx`) com base em um conjunto flex�vel de regras. O servi�o processa o arquivo Excel, aplica as regras de valida��o definidas para cada coluna e retorna o arquivo modificado, destacando as c�lulas inv�lidas com fundo rosa claro e coment�rios explicativos.

## Como funciona

- O servi�o recebe um arquivo Excel codificado em Base64 e uma lista de regras de valida��o.
- Cada regra � associada a uma coluna espec�fica do Excel.
- Para cada c�lula, todas as regras definidas para a respectiva coluna s�o aplicadas.
- Se uma c�lula n�o atender a uma ou mais regras, ela � destacada e recebe um coment�rio com as mensagens de erro.
- O arquivo Excel validado � retornado tamb�m em Base64.

## Exemplo de uso

### Estrutura da requisi��o

A requisi��o deve conter:

- `Base64Excel`: string com o arquivo Excel codificado em Base64.
- `Rules`: lista de regras de valida��o, onde cada regra possui:
  - `ColumnName`: nome da coluna a ser validada (deve coincidir com o cabe�alho do Excel).
  - `RuleType`: tipo da regra (ex: `Required`, `Regex`, `MaxLength`, `MinLength`, `Numeric`, etc).
  - `RuleValue`: valor de refer�ncia para a regra (quando aplic�vel).
  - `ErrorMessage`: mensagem de erro personalizada.

    #### Exemplo de defini��o de regras
   
 
var request = new ValidationRequest
{
    Base64Excel = "<arquivo_excel_em_base64>",
    Rules = new List<ValidationRule>
    {
        new ValidationRule
        {
            ColumnName = "Email",
            RuleType = "Required",
            ErrorMessage = "O e-mail � obrigat�rio."
        },
        new ValidationRule
        {
            ColumnName = "Email",
            RuleType = "Regex",
            RuleValue = @"^[\w.-]+@[\w.-]+\.\w+$",
            ErrorMessage = "Formato de e-mail inv�lido."
        },
        new ValidationRule
        {
            ColumnName = "Idade",
            RuleType = "NumericGTZero",
            ErrorMessage = "A idade deve ser um n�mero maior que zero."
        },
        new ValidationRule
        {
            ColumnName = "Nome",
            RuleType = "MaxLength",
            RuleValue = "50",
            ErrorMessage = "O nome deve ter no m�ximo 50 caracteres."
        }
    }
};

#### Observa��o:  
> Voc� pode definir m�ltiplas regras para a mesma coluna, e todas ser�o aplicadas. Caso uma c�lula viole mais de uma regra, todas as mensagens de erro ser�o exibidas no coment�rio da c�lula.

### Tipos de regras suportadas

- `Required`: valor obrigat�rio (n�o pode ser vazio).
- `Regex`: valida��o por express�o regular.
- `MaxLength`: tamanho m�ximo do texto.
- `MinLength`: tamanho m�nimo do texto.
- `Length`: tamanho exato do texto.
- `GTIN`: valida��o de c�digo GTIN/EAN.
- `Numeric`: valor num�rico.
- `NumericGTZero`: valor num�rico maior que zero.

## Como executar

1. Adicione a refer�ncia ao pacote [ClosedXML](https://github.com/ClosedXML/ClosedXML).
2. Importe e utilize o servi�o `ExcelValidationService` em seu projeto.
3. Envie a requisi��o conforme o exemplo acima.
4. O resultado ser� uma string Base64 do arquivo Excel validado, pronto para download ou exibi��o.

---

Sinta-se � vontade para adaptar as regras conforme a necessidade do seu dom�nio!