
using System.Collections.Generic;

namespace ExcelValidator.Models
{
    public class ValidationRequest
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ValidationRequest"/> class,  representing a request to validate
        /// an Excel file against a set of rules.
        /// </summary>
        /// <param name="FileName">The name of the Excel file to be validated. Cannot be null or empty.</param>
        /// <param name="Base64Excel">A Base64-encoded string representing the contents of the Excel file.  This must be a valid Base64 string and
        /// represent a valid Excel file.</param>
        /// <param name="Rules">A list of validation rules to apply to the Excel file.  Cannot be null, but can be an empty list if no rules
        /// are required.</param>
        public ValidationRequest( string FileName, string Base64Excel, List<ValidationRule> Rules) 
        {
            /// Validate parameters
            this.FileName = FileName;
            this.Base64Excel = Base64Excel;
            this.Rules = Rules;
        }
        /// <summary>
        /// Gets or sets the name of the file, including its extension.
        /// </summary>
        public string FileName { get; set; } = null!;
        /// <summary>
        /// Gets or sets the Base64-encoded string representation of an Excel file.
        /// </summary>
        public string Base64Excel { get; set; } = null!;
        /// <summary>
        /// Gets or sets the collection of validation rules to be applied.
        /// </summary>
        public List<ValidationRule> Rules { get; set; } = new();
    }
}
