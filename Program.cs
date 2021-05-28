using ExcelDataReader;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace ExcelSearcher
{
    class Program
    {
        public static string fileAPath = @"C:\Users\devpc\source\repos\ExcelSearcher\6U leverage missing URGENT 2nd batch.xlsb.xlsx";
        public static string fileBPath = @"C:\Users\devpc\source\repos\ExcelSearcher\ImageMasterfile.xlsx";
        public static string fileATargetSheetName = "Sheet4";
        public static string fileBTargetSheetName = "IMAGES";
        public static int fileAQueryColumnIndex = 0; // starting from 0
        public static int fileAQuery2ColumnIndex = 3;
        public static int fileAResultPopulateColumnIndex = 4;
        public static int fileBFirstQueryColumnIndex = 0;
        public static int fileBSecondQueryColumnIndex = 4;
        public static int fileBResultColumnIndex = 6;

        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            Console.WriteLine("File Searched From: ");
            fileAPath = Console.ReadLine();
            Console.WriteLine("File Main With the many rows: ");
            fileBPath = Console.ReadLine();
            IWorkbook fileAWorkBook;
            using (var fileAStream = new FileStream(fileAPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                try
                {
                    // Try to read workbook as XLSX:
                    try
                    {
                        fileAWorkBook = new XSSFWorkbook(fileAStream);
                    }
                    catch
                    {
                        fileAWorkBook = null;
                    }

                    // If reading fails, try to read workbook as XLS:
                    if (fileAWorkBook == null)
                    {
                        fileAWorkBook = new HSSFWorkbook(fileAStream);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Excel read error!");
                    return;
                }

                var targetSheet = fileAWorkBook.GetSheet(fileATargetSheetName);
                int rowIndex = 0;
                foreach(IRow row in targetSheet)
                {
                    var query1Cell = row.GetCell(fileAQueryColumnIndex);
                    var query2Cell = row.GetCell(fileAQuery2ColumnIndex);
                    var query1 = query1Cell.StringCellValue;
                    var query2 = query2Cell.StringCellValue;
                    var result = ReturnValueByQuery(query1, query2);
                    if (result == null)
                    {
                        Console.WriteLine("Not found for - " + query1 + " and " + query2);
                    }
                    else
                    {
                        WriteResult(fileAWorkBook, result, rowIndex);
                        Console.WriteLine("Result of searched value - " + query1 + " and " + query2 + " : " + result);
                    }

                    rowIndex++;
                } 
            }

            using (var fileAStream = new FileStream(fileAPath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                fileAWorkBook.Write(fileAStream);
            }
        }

        public static void WriteResult(IWorkbook workBook, string result, int rowIndex)
        {

            var targetSheet = workBook.GetSheet(fileATargetSheetName);
            var targetRow = targetSheet.GetRow(rowIndex);
            var targetCell = targetRow.CreateCell(fileAResultPopulateColumnIndex, CellType.String);
            targetCell.SetCellValue(result);
        }

        public static string ReturnValueByQuery(string query1, string query2)
        {
            using (var fileBStream = File.Open(fileBPath, FileMode.Open, FileAccess.Read))
            {
                using (var fileBReader = ExcelReaderFactory.CreateReader(fileBStream))
                {
                    do
                    {
                        if (fileBReader.Name == fileBTargetSheetName)
                        {
                            int rowIndex = 0;
                            while (fileBReader.Read())
                            {
                                if (rowIndex > 0)//skipping first row as its the title of the columns
                                {
                                    var firstColumnValueObj = fileBReader.GetValue(fileBFirstQueryColumnIndex);
                                    var secondColumnValueObj = fileBReader.GetValue(fileBSecondQueryColumnIndex);
                                    if (firstColumnValueObj != null && secondColumnValueObj != null)
                                    {
                                        string firstColumnToSearchValue = firstColumnValueObj.ToString();
                                        string secondColumnToSearchValue = secondColumnValueObj.ToString();
                                        bool foundInFirstColumn = firstColumnToSearchValue.Contains(query1, StringComparison.InvariantCultureIgnoreCase);
                                        if (foundInFirstColumn)
                                        {
                                            bool foundInSecondColumn = secondColumnToSearchValue.Contains(query2, StringComparison.InvariantCultureIgnoreCase);
                                            if (foundInSecondColumn)
                                            {
                                                var result = fileBReader.GetValue(fileBResultColumnIndex).ToString();
                                                return result;
                                            }
                                        }
                                    }
                                }

                                rowIndex++;
                            }
                        }
                    } while (fileBReader.NextResult());
                }
            }

            return null;
        }
    }
}
