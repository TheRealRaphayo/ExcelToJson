using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Newtonsoft.Json;
using ClosedXML.Excel;
using System.IO;

namespace ExcelToJson
{
    class Program
    {
        static void Main()
        {
            string excelPath = "language.xlsx"; // Update with your file path
            string jsonPath = "language.json";
            Console.WriteLine($"Will convert {excelPath} to {jsonPath}");

            /*
            var dictionary = ConvertExcelToJson(excelPath);
            string json = JsonConvert.SerializeObject(dictionary, Newtonsoft.Json.Formatting.Indented);
            */

            var dataList = ConvertExcelToList(excelPath);
            string json = JsonConvert.SerializeObject(dataList, Newtonsoft.Json.Formatting.Indented);

            File.WriteAllText(jsonPath, json, Encoding.GetEncoding("ISO-8859-1"));
            Console.WriteLine("JSON file created successfully with ISO-8859-1 encoding!");
            Console.WriteLine("Press any key to close");
            Console.ReadKey();
        }

        static List<DataEntry> ConvertExcelToList(string filePath)
        {
            var result = new List<DataEntry>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1); // Assuming data is in the first sheet
                var firstRow = worksheet.FirstRowUsed(); // Get the first row
                int maxColumns = firstRow.CellsUsed().Count(); // Determine last non-empty column

                foreach (var row in worksheet.RowsUsed().Skip(1)) // Skip header row
                {
                    string key = row.Cell(1).GetString().Trim(); // First column as key
                    var labels = new List<string>();

                    // Read from column 2 to the last detected non-empty column
                    for (int col = 2; col <= maxColumns; col++)
                    {
                        string label = row.Cell(col).GetString().Trim();
                        labels.Add(label); // Add label (including empty ones)
                    }

                    if (!string.IsNullOrEmpty(key))
                    {
                        result.Add(new DataEntry { Tag = key, Labels = labels });
                    }
                }
            }
            return result;
        }

        static Dictionary<string, List<string>> ConvertExcelToJson(string filePath)
        {
            var result = new Dictionary<string, List<string>>();
            try
            {

                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1); // Assuming data is in the first sheet
                    var firstRow = worksheet.FirstRowUsed(); // Get the first row
                    int maxColumns = firstRow.CellsUsed().Count(); // Determine last non-empty column

                    foreach (var row in worksheet.RowsUsed().Skip(1)) // Skip header row
                    {
                        string key = row.Cell(1).GetString().Trim(); // First column as key
                        if (!string.IsNullOrEmpty(key))
                        {
                            var labels = new List<string>();

                            // Read from column 2 to the last detected non-empty column
                            for (int col = 2; col <= maxColumns; col++)
                            {
                                string label = row.Cell(col).GetString().Trim();
                                labels.Add(label); // Add label, even if it's empty
                            }

                            result[key] = labels;
                        }
                    }
                }
               // return result;
            }
            catch (Exception e)
            {
                result["ExceptionError"] = new List<string>() { e.Message, e.StackTrace, e.InnerException.ToString() };
            }
            return result;
        }
    }

    // Class to represent each row in the JSON output
    class DataEntry
    {
        public string Tag { get; set; }
        public List<string> Labels { get; set; }
    }
}
