using ClosedXML.Excel;
using ExcelValidator.Models;
using System.Text.RegularExpressions;

namespace ExcelValidator.Services
{
    public class ExcelValidationService
    {
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
                foreach (var cell in row.Cells())
                {
                    var columnIndex = cell.Address.ColumnNumber;
                    var columnName = headers.ElementAtOrDefault(columnIndex - 1);

                    if (columnName == null || !ruleMap.TryGetValue(columnName, out var rule))
                        continue;

                    var value = cell.Value.ToString() ?? "";

                    if (!IsValid(value, rule))
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.LightPink;
                        cell.CreateComment().AddText(rule.ErrorMessage ?? "Invalid data");
                    }
                }
            }

            using var output = new MemoryStream();
            workbook.SaveAs(output);
            return Convert.ToBase64String(output.ToArray());
        }

        private bool IsValid(string value, ValidationRule rule)
        {
            return rule.RuleType switch
            {
                "Required" => !string.IsNullOrWhiteSpace(value),
                "Regex" => Regex.IsMatch(value, rule.RuleValue ?? ""),
                "MaxLength" => value.Length <= int.Parse(rule.RuleValue ?? "0"),
                _ => true,
            };
        }
    }
}
