using ClosedXML.Excel;

namespace LeituraArquivo.Controller
{
    public class LeituraController
    {
        private readonly string filePath = @"D:\STAF\CC_2024.xlsx";
        private const string worksheetName = "ELENCO_TCE_2024";

        public LeituraController()
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine("O arquivo de Excel não existe.");
                    return;
                }

                LerDados();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ocorreu um erro: {ex.Message}");
            }
        }

        private void LerDados()
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(worksheetName);
                var dataRows = worksheet.RowsUsed().Skip(1);

                if (!dataRows.Any())
                {
                    Console.WriteLine("Não há dados na planilha.");
                    return;
                }

                foreach (var row in dataRows)
                {
                    var (codigo, descricao, pai) = (row.Cell("I").GetString(), row.Cell("J").GetString(), row.Cell("P").GetString());
                    Console.WriteLine($"{codigo} - {descricao} - {pai}");
                }
            }
        }
    }
}