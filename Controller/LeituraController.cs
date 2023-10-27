using ClosedXML.Excel;

namespace LeituraArquivo.Controller
{
    public class LeituraController
    {
        public LeituraController()
        {

        }

        public async Task Ler()
        {
            var xls = new XLWorkbook(@"D:\STAF\CC_2024.xlsx");
            var planilha = xls.Worksheets.First(w => w.Name == "ELENCO_TCE_2024");
            var totalLinhas = planilha.Rows().Count();
            // primeira linha é o cabecalho
            for (int l = 2; l <= totalLinhas; l++)
            {
                var codigo = planilha.Cell($"I{l}").Value.ToString();
                var descricao = planilha.Cell($"J{l}").Value.ToString();
                var pai = planilha.Cell($"P{l}").Value.ToString();
                Console.WriteLine($"{codigo} - {descricao} - {pai}");
            }
        }
    }
}
