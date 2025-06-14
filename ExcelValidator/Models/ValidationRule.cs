
namespace ExcelValidator.Models
{
    public class ValidationRule
    {
        public string ColumnName { get; set; } = null!;
        public string RuleType { get; set; } = null!;
        public string? RuleValue { get; set; }
        public string? ErrorMessage { get; set; }
    }
}
