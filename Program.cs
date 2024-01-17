using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using DocumentFormat.OpenXml.Spreadsheet;



namespace CSharp___Gerador_de_Planilhas
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Lista para armazenar as despesas
            List<Despesa> despesas = new List<Despesa>();

            // Loop para inserção de despesas
            while (true)
            {
                // Solicita ao usuário o nome da despesa
                Console.Write("Informe o nome da despesa (ou 'sair' para encerrar): ");
                string nomeDespesa = Console.ReadLine();

                // Verifica se o usuário deseja encerrar
                if (nomeDespesa.ToLower() == "sair")
                    break;

                // Solicita ao usuário o valor da despesa
                Console.Write("Informe o valor da despesa: ");
                double valorDespesa;
                if (!double.TryParse(Console.ReadLine(), out valorDespesa))
                {
                    Console.WriteLine("Valor inválido. Tente novamente.");
                    continue;
                }

                // Adiciona a despesa à lista
                despesas.Add(new Despesa(nomeDespesa, valorDespesa));
            }

            // Exibe as despesas e totaliza os gastos
            double totalGastos = 0;
            Console.WriteLine("\nDespesas registradas:");
            foreach (var despesa in despesas)
            {
                Console.WriteLine($"{despesa.Nome}: {despesa.Valor:C}");
                totalGastos += despesa.Valor;
            }

            // Exibe o total de gastos
            Console.WriteLine($"\nTotal de Gastos: {totalGastos:C}");

            // Calcula e exibe a distribuição percentual
            Console.WriteLine("\nDistribuição Percentual:");
            foreach (var despesa in despesas)
            {
                double percentual = (despesa.Valor / totalGastos) * 100;
                Console.WriteLine($"{despesa.Nome}: {percentual:F2}%");
            }

            // Gera a planilha no Excel
            GerarPlanilha(despesas);

            Console.ReadLine(); // Aguarda uma tecla ser pressionada antes de encerrar
        }

        static void GerarPlanilha(List<Despesa> despesas)
        {
            // Criar um novo workbook do NPOI
            IWorkbook workbook = new XSSFWorkbook();

            // Criar uma nova planilha no workbook
            ISheet sheet = workbook.CreateSheet("Despesas");

            // Cabeçalhos
            IRow headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("Nome");
            headerRow.CreateCell(1).SetCellValue("Valor");

            // Preencher os dados
            for (int i = 0; i < despesas.Count; i++)
            {
                IRow dataRow = sheet.CreateRow(i + 1);
                dataRow.CreateCell(0).SetCellValue(despesas[i].Nome);
                dataRow.CreateCell(1).SetCellValue(despesas[i].Valor);
            }

            // Salvar o arquivo Excel
            using (var fileStream = new FileStream("C:\\Users\\USUARIO\\Desktop\\C#\\Gerador De planinha\\Gerador de planinha\\planinhas\\Despesas_NPOI.xlsx", FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }

            Console.WriteLine("Planilha gerada com sucesso (Despesas_NPOI.xlsx)");
        }
    }

    class Despesa
    {
        public string Nome { get; }
        public double Valor { get; }

        public Despesa(string nome, double valor)
        {
            Nome = nome;
            Valor = valor;
        }
    }
}