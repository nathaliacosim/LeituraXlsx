using LeituraArquivo.Controller;

public class Program
{
    private static async Task Main(string[] args)
    {
        LeituraController leituraController = new LeituraController();
        await leituraController.Ler();
    }
}