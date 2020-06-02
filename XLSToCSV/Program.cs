using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDataReader;

namespace XLSToCSV
{
    class Program
    {
        private static NumberFormatInfo customNumFormat;
        
        static void Main(string[] args)
        {
            Console.WriteLine("Starting console...");
            if (!Directory.Exists($"{Environment.CurrentDirectory}/Input"))
                Directory.CreateDirectory($"{Environment.CurrentDirectory}/Input");
            if (!Directory.Exists($"{Environment.CurrentDirectory}/Output"))
                Directory.CreateDirectory($"{Environment.CurrentDirectory}/Output");
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            customNumFormat = (NumberFormatInfo) CultureInfo.InvariantCulture.NumberFormat.Clone();
            customNumFormat.NumberDecimalSeparator = ",";
            customNumFormat.NumberGroupSeparator = "";
            
            var files = Directory.GetFiles($"{Environment.CurrentDirectory}/Input").ToList();
            for (int i = 0; i < files.Count; i++)
            {
                if (!files[i].ToLower().Contains("xls"))
                {
                    Console.WriteLine($"File is not Excel file: {Path.GetFileName(files[i])}");
                    files.RemoveAt(i);
                }
            }
            Console.WriteLine($"Found {files.Count} Excel files in input folder");
            
            foreach (var file in files)
            {
                var currentFile = file;
                currentFile = currentFile.Replace("$", "");
                Console.WriteLine($"Processing file ${Path.GetFileName(currentFile)}");
                var fileName = Path.GetFileNameWithoutExtension(currentFile) + ".csv";
                var saved = SaveAsCsv(currentFile, $"{Environment.CurrentDirectory}/Output/{fileName}");
                if (!saved)
                    Console.WriteLine($"Couldn't convert {fileName} to CSV");
            }

            Console.WriteLine("Processing done");
            Console.WriteLine("Press ENTER to close the window...");
        }
        
        public static bool SaveAsCsv(string excelFilePath, string destinationCsvFilePath)
        {

            using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IExcelDataReader reader = null;
                if (excelFilePath.EndsWith(".xls"))
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (excelFilePath.EndsWith(".xlsx"))
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                if (reader == null)
                    return false;

                var ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = false
                    }
                });

                for (int tableCounter = 0; tableCounter < ds.Tables.Count; tableCounter++)
                {
                    var csvContent = string.Empty;
                    int row_no = 0;
                    while (row_no < ds.Tables[tableCounter].Rows.Count)
                    {
                        var percentage = ((double) row_no / (double) ds.Tables[tableCounter].Rows.Count) * 100;
                        percentage = Math.Round(percentage, 2);
                        Console.Write($"Processing row #{row_no} / {ds.Tables[tableCounter].Rows.Count} ({percentage}%)\r");
                        var arr = new List<string>();
                        for (int i = 0; i < ds.Tables[tableCounter].Columns.Count; i++)
                        {
                            var currArr = ds.Tables[tableCounter].Rows[row_no][i];
                            if(!double.TryParse(currArr.ToString(), out var currArrDeci))
                                arr.Add(ds.Tables[tableCounter].Rows[row_no][i].ToString());
                            else if (currArrDeci % 1 != 0)
                                arr.Add(currArrDeci.ToString("N", customNumFormat));
                            else 
                                arr.Add(ds.Tables[tableCounter].Rows[row_no][i].ToString());
                        }
                        row_no++;
                        csvContent += string.Join(";", arr) + "\n";
                    }

                    var fileName = Path.GetFileNameWithoutExtension(excelFilePath);
                    fileName = (ds.Tables[tableCounter].TableName) + "_" + fileName + ".csv";
                    var destinationPath = Path.GetDirectoryName(destinationCsvFilePath);
                    destinationCsvFilePath = destinationPath + "/" + fileName;
                    StreamWriter csv = new StreamWriter(destinationCsvFilePath, false);
                    csv.Write(csvContent);
                    csv.Close();
                    Console.WriteLine($"CSV Saved: /Output/{Path.GetFileName(destinationCsvFilePath)}");
                }
                
                return true;
            }
        }
    }
}