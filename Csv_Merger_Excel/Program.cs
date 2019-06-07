using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Microsoft.VisualBasic.FileIO;

namespace Csv_Merger_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            var files = Directory.GetFiles(@"C:\Data\AzureDataStudio\Pokeapi\data\v2\csv");

            using (var excel = new ExcelPackage())
            {
                foreach (var item in files)
                {
                    var start = item.IndexOf(@"v2\csv\") + 7;
                    var end = item.IndexOf(".csv");
                    var filename = item.Substring(start, end - start);

                    Console.WriteLine(filename);

                    excel.Workbook.Worksheets.Add(filename);

                    var currentWorksheet = excel.Workbook.Worksheets[filename];

                    using (TextFieldParser csvParser = new TextFieldParser(item))
                    {
                        Console.WriteLine("inner loop");

                        csvParser.CommentTokens = new string[] { "#" };
                        csvParser.SetDelimiters(new string[] { "," });
                        csvParser.HasFieldsEnclosedInQuotes = true;

                        var cellData = new List<string[]>();

                        while (!csvParser.EndOfData)
                        {
                            // Read current line fields, pointer moves to the next line.
                            cellData.Add(csvParser.ReadFields());
                        }
                        currentWorksheet.Cells[1, 1].LoadFromArrays(cellData);
                    }
                }
                FileInfo excelFile = new FileInfo(@"C:\Data\AzureDataStudio\Pokeapi\data\v2\csv\output\neatwb.xlsx");
                excel.SaveAs(excelFile);
            }
        }
    }
}
