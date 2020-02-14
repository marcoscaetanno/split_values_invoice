using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entity
{
    public class CsvRules
    {
        /// <summary>
        /// Colunas do arquivo CSV
        /// </summary>
        public class CsvEntity
        {
            public string Titulo { get; set; }
            public string Categoria { get; set; }
            public string Data { get; set; }
            public string Valor { get; set; }
        }

        public class CsvInfos 
        {
            public string Diretorio { get; set; }
            public string NomeArquivo { get; set; }
        }
    }
}
