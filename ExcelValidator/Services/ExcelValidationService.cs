using ClosedXML.Excel;
using ExcelValidator.Models;
using System.Text.RegularExpressions;

namespace ExcelValidator.Services
{
    /// <summary>
    /// Fornece funcionalidades para validar um arquivo Excel com base em um conjunto de regras e retornar o arquivo modificado com os resultados da validação.
    /// </summary>
    /// <remarks>
    /// Este serviço processa um arquivo Excel fornecido como uma string codificada em Base64, valida seu conteúdo com base nas regras especificadas e destaca as células inválidas no arquivo Excel resultante. Células inválidas são marcadas com fundo rosa claro e um comentário contendo a mensagem de erro. O arquivo Excel modificado é retornado como uma string codificada em Base64.
    /// </remarks>
    public class ExcelValidationService
    {
        /// <summary>
        /// Valida o conteúdo de um arquivo Excel com base nas regras de validação especificadas e retorna o arquivo modificado como uma string codificada em Base64.
        /// </summary>
        /// <remarks>
        /// O método processa a primeira planilha do arquivo Excel fornecido. Cada célula é validada conforme as regras especificadas no <paramref name="request"/>. Se o valor de uma célula não atender à regra correspondente, a célula é destacada com fundo rosa claro e um comentário é adicionado com a mensagem de erro.
        /// </remarks>
        /// <param name="request">A requisição de validação contendo o arquivo Excel codificado em Base64 e o conjunto de regras de validação a serem aplicadas.</param>
        /// <returns>
        /// Uma string codificada em Base64 representando o arquivo Excel modificado. Células com dados inválidos são destacadas e comentários são adicionados para indicar os erros de validação.
        /// </returns>
        public string ValidateExcel(ValidationRequest request)
        {
            using var stream = new MemoryStream(Convert.FromBase64String(request.Base64Excel));
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheet(1);

            var headerRow = worksheet.Row(1);
            var headers = headerRow.Cells().Select(c => c.Value.ToString()).ToList();

            var ruleMap = request.Rules.ToDictionary(r => r.ColumnName, r => r);

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                for (int colIndex = 1; colIndex <= headers.Count; colIndex++)
                {
                    var columnName = headers.ElementAtOrDefault(colIndex - 1);
                    if (columnName == null || !ruleMap.TryGetValue(columnName, out var rule))
                        continue;

                    var cell = row.Cell(colIndex);
                    var value = cell.Value.ToString() ?? "";

                    if (!IsValid(value, rule))
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.LightPink;
                        cell.CreateComment().AddText(rule.ErrorMessage ?? "Dado inválido");
                    }
                }
            }

            using var output = new MemoryStream();
            workbook.SaveAs(output);
            return Convert.ToBase64String(output.ToArray());
        }

        /// <summary>
        /// Valida um valor de string com base em uma regra de validação especificada.
        /// </summary>
        /// <remarks>
        /// O método suporta várias regras de validação, como verificação de obrigatoriedade, correspondência de expressões regulares, restrições de tamanho, validação de valores numéricos e outras. O comportamento da validação depende do <see cref="ValidationRule.RuleType"/> e do respectivo <see cref="ValidationRule.RuleValue"/>.
        /// </remarks>
        /// <param name="value">O valor em string a ser validado. Pode ser nulo ou vazio dependendo da regra aplicada.</param>
        /// <param name="rule">A <see cref="ValidationRule"/> que define o tipo de validação a ser realizada e quaisquer parâmetros associados.</param>
        /// <returns>
        /// <see langword="true"/> se o <paramref name="value"/> satisfaz a <paramref name="rule"/> especificada; caso contrário, <see langword="false"/>.
        /// </returns>
        private bool IsValid(string value, ValidationRule rule)
        {
            return rule.RuleType switch
            {
                "Required" => !string.IsNullOrWhiteSpace(value),
                "Regex" => Regex.IsMatch(value, rule.RuleValue ?? ""),
                "MaxLength" => value.Length <= int.Parse(rule.RuleValue ?? "0"),
                "Length" => value.Length == int.Parse(rule.RuleValue ?? "0"),
                "MinLength" => value.Length >= int.Parse(rule.RuleValue ?? "0"),
                "GTIN" => value.IsValidEan(),
                "Numeric" => double.TryParse(value, out _),
                "NumericGTZero" => double.TryParse(value, out var n) && n > 0,
                _ => true,
            };
        }
    }
}
