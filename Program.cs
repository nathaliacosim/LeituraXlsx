using LeituraArquivo.Controller;
using System;
using System.IO;
using System.Linq;

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            var leituraController = new LeituraController();
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"Erro: {ex.Message}. Certifique-se de que o arquivo Excel existe no caminho especificado.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ocorreu um erro inesperado: {ex.Message}");
        }
        finally
        {
            Console.WriteLine("O programa foi encerrado.");
        }
    }
}