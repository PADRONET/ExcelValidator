using ExcelValidator.Models;
using ExcelValidator.Services;
using ClosedXML.Excel;

namespace TesteExcelValidator
{
    public class UnitTest1
    {
       [Fact]
        public void ValidateExcel_DeveDestacarCelulaInvalida()
        {
            // Arrange: cria um arquivo Excel em memória
            using var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Plan1");
            worksheet.Cell(1, 1).Value = "Nome";
            worksheet.Cell(2, 1).Value = ""; // Valor inválido para "Required"
            worksheet.Cell(1, 2).Value = "SobreNome";
            worksheet.Cell(2, 2).Value = "Padrão"; // Valor valido para "Required"
            worksheet.Cell(1, 3).Value = "CODIGO_EAN";
            worksheet.Cell(2, 3).Value = "0"; // Valor invalido para "GTIN"
            worksheet.Cell(1, 4).Value = "VALOR";
            worksheet.Cell(2, 4).Value = "2"; // Valor valido para "Numeric"
            worksheet.Cell(1, 5).Value = "VALOR2";
            worksheet.Cell(2, 5).Value = ""; // Valor invalido para "Numeric"
            worksheet.Cell(1, 6).Value = "NCM";
            worksheet.Cell(2, 6).Value = "99999999"; // Valor valido para "Length"
            worksheet.Cell(1, 7).Value = "CEST";
            worksheet.Cell(2, 7).Value = "99999"; // Valor invalido para "MaxLength"

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            var base64Excel = Convert.ToBase64String(ms.ToArray());

            var regras = new List<ValidationRule>
            {
                new ValidationRule("Nome", "Required", null, "Nome é obrigatório"),
                new ValidationRule("SobreNome", "Required", null, "SobreNome é obrigatório"),
                new ValidationRule("CODIGO_EAN", "GTIN", null, "Codigo EAN Invalido"),
                new ValidationRule("VALOR", "Numeric", null, "Não é um numero"),
                new ValidationRule("VALOR2", "Numeric", null, "Não é um numero"),
                new ValidationRule("NCM", "Length", "8", "Formato invalido"),
                new ValidationRule("CEST", "MaxLength", "4", "Formato invalido")
            };

            var request = new ValidationRequest("teste.xlsx", base64Excel, regras);
            var service = new ExcelValidationService();

            // Act
            var resultadoBase64 = service.ValidateExcel(request);

            // Assert: verifica se o resultado contém alguma alteração (cor de fundo ou comentário)
            var resultadoBytes = Convert.FromBase64String(resultadoBase64);
            using var resultadoStream = new MemoryStream(resultadoBytes);
            using var resultadoWorkbook = new XLWorkbook(resultadoStream);
            var resultadoWorksheet = resultadoWorkbook.Worksheet(1);
            var cell = resultadoWorksheet.Cell(2, 1);
            var cell2 = resultadoWorksheet.Cell(2, 2);
            var cell3 = resultadoWorksheet.Cell(2, 3);
            var cell4 = resultadoWorksheet.Cell(2, 4);
            var cell5 = resultadoWorksheet.Cell(2, 5);
            var cell6 = resultadoWorksheet.Cell(2, 6);
            var cell7 = resultadoWorksheet.Cell(2, 7);

            // A célula deve estar com fundo rosa claro (LightPink)
            Assert.Equal(ClosedXML.Excel.XLColor.LightPink, cell.Style.Fill.BackgroundColor);
            Assert.Contains("Nome é obrigatório", cell.GetComment().Text);

            // A célula nao deve estar com fundo rosa claro (LightPink)
            Assert.NotEqual(ClosedXML.Excel.XLColor.LightPink, cell2.Style.Fill.BackgroundColor);
            Assert.DoesNotContain("Nome é obrigatório", cell2.GetComment()?.Text ?? "");

            // A célula nao deve estar com fundo rosa claro (LightPink)
            Assert.Equal(ClosedXML.Excel.XLColor.LightPink, cell3.Style.Fill.BackgroundColor);
            Assert.Contains("Codigo EAN Invalido", cell3.GetComment()?.Text ?? "");

            // A célula nao deve estar com fundo rosa claro (LightPink)
            Assert.NotEqual(ClosedXML.Excel.XLColor.LightPink, cell4.Style.Fill.BackgroundColor);
            Assert.DoesNotContain("Não é um numero", cell4.GetComment()?.Text ?? "");

            // A célula nao deve estar com fundo rosa claro (LightPink)
            Assert.Equal(ClosedXML.Excel.XLColor.LightPink, cell5.Style.Fill.BackgroundColor);
            Assert.Contains("Não é um numero", cell5.GetComment()?.Text ?? "");

            // A célula nao deve estar com fundo rosa claro (LightPink)
            Assert.NotEqual(ClosedXML.Excel.XLColor.LightPink, cell6.Style.Fill.BackgroundColor);
            Assert.DoesNotContain("Formato invalido", cell6.GetComment()?.Text ?? "");

            // A célula nao deve estar com fundo rosa claro (LightPink)
            Assert.Equal(ClosedXML.Excel.XLColor.LightPink, cell7.Style.Fill.BackgroundColor);
            Assert.Contains("Formato invalido", cell7.GetComment()?.Text ?? "");
        }
    }
}
