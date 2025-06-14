
namespace ExcelValidator.Models
{
    /// <summary>
    /// Represents a validation rule that can be applied to a specific column in a data set.
    /// </summary>
    /// <remarks>A validation rule defines the criteria that a column's value must meet, along with an
    /// optional error message to display when the rule is violated. The rule is defined by specifying the column name,
    /// the type of rule, an optional rule value, and an optional error message.</remarks>
    public class ValidationRule
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ValidationRule"/> class, representing a rule for validating a
        /// specific column.
        /// </summary>
        /// <param name="ColumnName">The name of the column to which the validation rule applies. Cannot be null or empty.</param>
        /// <param name="RuleType">The type of validation to be performed (e.g., "Required", "MaxLength", "Regex"). Cannot be null or empty.</param>
        /// <param name="RuleValue">The value associated with the validation rule, such as a maximum length or a regular expression pattern. Can
        /// be null if the rule type does not require a value.</param>
        /// <param name="ErrorMessage">The error message to display if the validation rule is violated. Can be null if no custom error message is
        /// required.</param>
        public ValidationRule( string ColumnName, string RuleType, string? RuleValue, string? ErrorMessage) 
        {
            // Validate parameters
            
            this.ColumnName = ColumnName;
            this.RuleType = RuleType;
            this.RuleValue = RuleValue;
            this.ErrorMessage = ErrorMessage;
        }
        /// <summary>
        /// Gets or sets the name of the database column associated with this entity.
        /// </summary>
        public string ColumnName { get; set; } = null!;
        /// <summary>
        /// Gets or sets the type of the rule represented as a string.
        /// </summary>
        public string RuleType { get; set; } = null!;
        /// <summary>
        /// Gets or sets the value associated with the rule.
        /// </summary>
        public string? RuleValue { get; set; }
        /// <summary>
        /// Gets or sets the error message associated with the current operation or state.
        /// </summary>
        public string? ErrorMessage { get; set; }
    }
}
