using System;
using System.Collections.Generic;
using ClosedXML.Excel; //Importando a biblioteca ClosedXML.
using Gerenciadorxml.Entitties;

namespace Gerenciadorxml
{
    class Program
    {
        static void Main(string[] args)
        {
           
            using (var workbook = new XLWorkbook(@"C:\Users\Renato Martins\Desktop\planilha\Produtividade 2019-2020.xlsx"))
            {

         

                var listaValores = new List<DadosExcel>();

                var contador = 1;
                var parada = false;
                var sheetParada = false;
                var linha = 2;

                while (!sheetParada)
                {
                    var planilha = workbook.Worksheet(contador);

                    if (planilha.Name.StartsWith("T"))
                    {
                     
                        Console.WriteLine("Sheet: " + planilha.Name);
                    
                        while (!parada)
                        {
                           

                            for (int coluna = 1; coluna < 28; coluna++)
                            {
                                var DadosExcel = new DadosExcel() { NomeSheet = planilha.Name, linha = linha, coluna = coluna, valor = planilha.Cell(linha, coluna).Value.ToString() };
                                listaValores.Add(DadosExcel);
                            }
                            linha++;

                            if (string.IsNullOrEmpty(planilha.Cell(linha, 1).Value.ToString()))
                            {
                                parada = true;
                            }

                        }
                    }
                    else
                    {
                        sheetParada = true;
                    }

                    parada = false;
                    linha = 2;
                    contador += 1;
                }
                workbook.Dispose();
            }
            Console.ReadKey();
        }
    }
}
