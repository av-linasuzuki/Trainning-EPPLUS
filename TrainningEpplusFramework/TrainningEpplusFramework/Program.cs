using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;

namespace TrainningEpplusFramework
{
    class Program
    {
        static void Main(string[] args)
        {
            // armazena o local que está planilha (não necessário inicialmente)
            var path = @"\CadastrarProdutos.xlsx";

            //para acessar uma planilha, criamos um pacote com o método ExcelPackage e acessamos com FileInfo
            var package = new ExcelPackage(new FileInfo(path));

            //para criar uma pasta de trabalho, criamos um pacote com o método ExcelPackage
            //var package = new ExcelPackage();

            /* para acessar a primeira planilha da pasta de trabalho, é preciso acessar o Workbook para então, 
             * acessar as planilhas da pasta de trabalho */

            // nesse caso, acessamos a planilha através do índice
            var firstSheetIndex = package.Workbook.Worksheets[0];
            Console.WriteLine("Nome da planilha recuperada: {0}", firstSheetIndex); // imprime o nome da planilha recuperada

            // recuperando o nome da primeira planilha através do nome
            var firstSheetName = package.Workbook.Worksheets["Cadastrar"];
            Console.WriteLine("Nome da planilha: {0}", firstSheetName);

            // recuperando o nome da primeira planilha através do método First
            var SheetName = package.Workbook.Worksheets.First();
            Console.WriteLine("Nome da planilha: {0}", SheetName);

            // acessar o valor de uma célula específica
            var valorCel = firstSheetName.GetValue(2, 4);
            Console.WriteLine("Valor recuperado da linha 2 coluna 4: {0}", valorCel);

            // adicionando uma nova planilha
            var newSheet = package.Workbook.Worksheets.Add("NovaPlanilha");

            // adicionando um novo valor em uma célula
            newSheet.Cells[1, 1].Value = "Meu novo valor";
            //formatando a célula
            newSheet.Cells[1, 1].Style.Font.Bold = true;

            // Salvando uma copia da pasta de trabalho
            package.SaveAs(new FileInfo(@"\CopiaPastaTrabalho.xlsx"));

        }
    }
}
