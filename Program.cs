using ExcelDataReader;
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
            using (var fileAStream = File.Open(fileAPath, FileMode.Open, FileAccess.Read))
            {
                using (var fileAReader = ExcelReaderFactory.CreateReader(fileAStream))
                {
                    do
                    {
                        if(fileAReader.Name == fileATargetSheetName)
                        {
                            int rowIndex = 0;
                            while (fileAReader.Read())
                            {
                                var query1 = fileAReader.GetValue(fileAQueryColumnIndex).ToString();
                                var query2Obj = fileAReader.GetValue(fileAQuery2ColumnIndex);
                                var query2 = query2Obj.ToString();
                                var result = ReturnValueByQuery(query1, query2);
                                if(result == null)
                                {
                                    Console.WriteLine("Not found for - " + query1 + " and " + query2);
                                } else
                                {
                                    Console.WriteLine("Result of searched value - " + query1 + " and " + query2 + " : " + result);
                                }

                                rowIndex++;
                            }
                        }
                    } while (fileAReader.NextResult());
                }
            }
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
