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
            string excelPath = "data.xlsx"; // Update with your file path
            string jsonPath = "output.json";

            var dictionary = ConvertExcelToJson(excelPath);
            string json = JsonConvert.SerializeObject(dictionary, Newtonsoft.Json.Formatting.Indented);

            File.WriteAllText(jsonPath, json);
            Console.WriteLine("JSON file created successfully!");
        }

        static Dictionary<string, List<string>> ConvertExcelToJson(string filePath)
        {
            var result = new Dictionary<string, List<string>>();

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
            return result;
        }
    }
}
