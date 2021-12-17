using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;

namespace Excel_Txt
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                ShowHelp();
                return;
            }

            string source = args[0];
            string dest = "";
            if (args.Length >1)
            {
                dest = args[1];
            }
               

            if (File.Exists(source))
                ConvertFile(source, dest);
            else if (Directory.Exists(source))
                ConvertDirectory(source, dest);
            else
                ShowHelp();
        }

        static void ConvertDirectory(string source, string destination)
        {
            string[] sourceFiles = Directory.GetFiles(source, "*", SearchOption.AllDirectories).Where(f => f.EndsWith(".xlsx")).ToArray(); ;
            foreach (string file in sourceFiles)
            {
                string target = file;
                if (destination != "") target = file.Replace(source, destination);
                ConvertFile(file, target);
            }
        }

        static void ConvertFile(string source, string destination)
        {
            if (destination == "")
                destination = source;

            destination = destination.Replace(".xlsx", ".txt");

            try
            {
                if (XlsxToTxt(source, destination))
                {
                    Console.WriteLine(destination + " [Success]");
                }
                else
                {
                    Console.WriteLine(destination + " [FAILED]");
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(destination + " [FAILED]");
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }

            txtsb.Clear();
        }

        static StringBuilder txtsb = new StringBuilder();
        static bool XlsxToTxt(string xlsx_file, string txt_file)
        {
            using (var excel = new ExcelPackage(new FileInfo(xlsx_file)))
            {
                var worksheets = excel.Workbook.Worksheets[1];

                if (worksheets.Tables.Count <=0)
                {
                    return false;
                }
                int columnNum = worksheets.Tables[0].Address.Columns;
                int rowNum = worksheets.Tables[0].Address.Rows;

                if (columnNum <=0 && rowNum <=0)
                {
                    return false;
                }
                for (int rI = 1; rI <= rowNum; rI++)
                {
                    for (int cI = 1; cI <= columnNum; cI++)
                    {
                        var value = worksheets.Cells[rI, cI].Value;
                        if (value != null)
                        {
                            txtsb.Append(value);
                        }
                        if (cI != columnNum)
                        {
                            txtsb.Append("\t");
                        }
                    }

                    if (rI != rowNum)
                    {
                        txtsb.Append("\n");
                    }
                }
                File.WriteAllText(txt_file, txtsb.ToString());

                txtsb.Clear();
            }
            return true;
        }

        static void ShowHelp()
        {
            Console.WriteLine(@"
    Excel To Txt:Only tables in Excel can be converted
    Excel-Txt [source] [destination] 

    source      : input file or directory
    destination : output file or directory, Same as the source if not specified");
        }

    }
}
