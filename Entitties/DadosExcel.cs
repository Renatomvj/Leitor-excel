using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.Excel; //Biblioteca importada.

namespace Gerenciadorxml.Entitties
{
    public class DadosExcel
    {
        public string NomeSheet { get; set; }
        public int linha { get; set; }
        public int coluna { get; set; }
        public string valor { get; set; }
    }
    
}
