using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using static Entity.CsvRules;

namespace SplitValues
{
    public class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Hello user");

                List<CsvInfos> path = OpenDirectory();

                foreach (var item in path)
                {
                    var csv = ReadCSV(item);

                    bool excel = WriteCsv(csv, item);

                    if (excel)
                    {
                        Console.WriteLine("I performed my function successfully! You can check the new invoice.");
                        Thread.Sleep(1000);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unfortunately I presented a problem in my execution, could you take a look?");
                Console.WriteLine("####################################################################");
                Console.WriteLine(ex.StackTrace.ToString());
                Console.WriteLine("####################################################################");
                Console.ReadKey();
                throw;
            }
        }

        public static List<CsvInfos> OpenDirectory()
        {
            try
            {
                List<CsvInfos> lstInfos = new List<CsvInfos>();

                List<string> files = new List<string>();
                string path = ConfigurationManager.AppSettings["CsvDirectory"];

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                    return lstInfos;
                }

                files = Directory.GetFiles(path).ToList();

                foreach (var item in files)
                {
                    CsvInfos info = new CsvInfos();
                    info.Diretorio = item;
                    info.NomeArquivo = item.Substring(17, 7);
                    lstInfos.Add(info);
                }

                return lstInfos;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Método responsável por ler as linhas do CSV e criar a lista com as linhas lidas
        /// </summary>
        /// <returns></returns>
        public static List<CsvEntity> ReadCSV(CsvInfos path)
        {
            try
            {
                Console.WriteLine("I will start reading your invoice.");

                List<CsvEntity> lstCsv = new List<CsvEntity>();

                using (var reader = new StreamReader(path.Diretorio))
                {
                    while (!reader.EndOfStream)
                    {
                        CsvEntity csv = new CsvEntity();

                        var lines = reader.ReadLine();

                        var values = lines.Split(',');

                        csv.Data = values[0];
                        csv.Categoria = values[1];
                        csv.Titulo = values[2];
                        csv.Valor = values[3];

                        lstCsv.Add(csv);
                    }
                }
                return lstCsv;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Método responsável por popular o arquivo Excel com lista de linhas do arquivo CSV
        /// </summary>
        /// <param name="csv"></param>
        public static bool WriteCsv(List<CsvEntity> csv, CsvInfos info)
        {
            try
            {
                Console.WriteLine("Just a moment I'll generate a new invoice ...");
                var dataTable = GetTable(csv);

                XLWorkbook wb = new XLWorkbook();

                var ws = wb.Worksheets.Add($@"Fatura {DateTime.Now.Month}");

                ws.Cell(1, 1).InsertData(dataTable.AsEnumerable());

                wb.SaveAs($@"C:\Fatura\Nova_Fatura\fatura_nubank_{info.NomeArquivo}.xlsx");

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static DataTable GetTable(List<CsvEntity> csv)
        {
            try
            {
                DataTable dt = new DataTable();

                dt.Columns.AddRange(new DataColumn[4] {
                    new DataColumn("Data", typeof(string)),
                    new DataColumn("Categoria", typeof(string)),
                    new DataColumn("Título", typeof(string)),
                    new DataColumn("Valor")});

                foreach (var line in csv)
                {
                    if (line.Data != "date")
                        Convert.ToDateTime(line.Data);

                    if (line.Valor != "amount")
                        Convert.ToDouble(line.Valor);

                    dt.Rows.Add(line.Data, line.Categoria, line.Titulo, line.Valor);
                }

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
