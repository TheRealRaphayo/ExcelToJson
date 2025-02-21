using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace ExcelToJsonApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.Text = "Convert Excel to Json";
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        string TxtFilePath;
        string JsonFilePath;

        private void BtnSelectFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //TxtFilePath.Text = openFileDialog.FileName;
                    this.TxtFilePath = openFileDialog.FileName;
                    this.JsonFilePath = Path.ChangeExtension(this.TxtFilePath, ".json");
                    BtnExportJson.Text = "Export " + openFileDialog.SafeFileName + " to JSON";
                    TextFilePath.Text = openFileDialog.FileName;
                }
            }
        }

        private void TextFilePath_TextChanged(object sender, EventArgs e)
        {

        }

        private void BtnExportJson_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TextFilePath.Text))
            {
                MessageBox.Show("Please select an Excel file first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try 
            {
                var dataList = ConvertExcelToList(this.TxtFilePath);
                string json = JsonConvert.SerializeObject(dataList, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(this.JsonFilePath, json, Encoding.GetEncoding("ISO-8859-1"));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            MessageBox.Show($"File created and save at {this.JsonFilePath}", "CoffeeTime", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        // Class to represent each row in the JSON output
        class DataEntry
        {
            public string Tag { get; set; }
            public List<string> Labels { get; set; }
        }
    }
}
