using ExcelValidator.Models;

namespace ExcelValidator.Interfaces
{
    public interface IExcelValidationService
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
        public Task<string> ValidateExcel(ValidationRequest request);
    }
}
