using ClosedXML.Excel;
using ExcelValidator.Models;
using System.Text.RegularExpressions;

namespace ExcelValidator.Services
{
    /// <summary>
    /// Fornece funcionalidades para validar um arquivo Excel com base em um conjunto de regras e retornar o arquivo modificado com os resultados da valida��o.
    /// </summary>
    /// <remarks>
    /// Este servi�o processa um arquivo Excel fornecido como uma string codificada em Base64, valida seu conte�do com base nas regras especificadas e destaca as c�lulas inv�lidas no arquivo Excel resultante. C�lulas inv�lidas s�o marcadas com fundo rosa claro e um coment�rio contendo a mensagem de erro. O arquivo Excel modificado � retornado como uma string codificada em Base64.
    /// </remarks>
    public class ExcelValidationService
    {
        /// <summary>
        /// Valida o conte�do de um arquivo Excel com base nas regras de valida��o especificadas e retorna o arquivo modificado como uma string codificada em Base64.
        /// </summary>
        /// <remarks>
        /// O m�todo processa a primeira planilha do arquivo Excel fornecido. Cada c�lula � validada conforme as regras especificadas no <paramref name="request"/>. Se o valor de uma c�lula n�o atender � regra correspondente, a c�lula � destacada com fundo rosa claro e um coment�rio � adicionado com a mensagem de erro.
        /// </remarks>
        /// <param name="request">A requisi��o de valida��o contendo o arquivo Excel codificado em Base64 e o conjunto de regras de valida��o a serem aplicadas.</param>
        /// <returns>
        /// Uma string codificada em Base64 representando o arquivo Excel modificado. C�lulas com dados inv�lidos s�o destacadas e coment�rios s�o adicionados para indicar os erros de valida��o.
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
                        cell.CreateComment().AddText(rule.ErrorMessage ?? "Dado inv�lido");
                    }
                }
            }

            using var output = new MemoryStream();
            workbook.SaveAs(output);
            return Convert.ToBase64String(output.ToArray());
        }

        /// <summary>
        /// Valida um valor de string com base em uma regra de valida��o especificada.
        /// </summary>
        /// <remarks>
        /// O m�todo suporta v�rias regras de valida��o, como verifica��o de obrigatoriedade, correspond�ncia de express�es regulares, restri��es de tamanho, valida��o de valores num�ricos e outras. O comportamento da valida��o depende do <see cref="ValidationRule.RuleType"/> e do respectivo <see cref="ValidationRule.RuleValue"/>.
        /// </remarks>
        /// <param name="value">O valor em string a ser validado. Pode ser nulo ou vazio dependendo da regra aplicada.</param>
        /// <param name="rule">A <see cref="ValidationRule"/> que define o tipo de valida��o a ser realizada e quaisquer par�metros associados.</param>
        /// <returns>
        /// <see langword="true"/> se o <paramref name="value"/> satisfaz a <paramref name="rule"/> especificada; caso contr�rio, <see langword="false"/>.
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
