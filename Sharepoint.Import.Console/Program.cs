using System;
using Microsoft.SharePoint.Client;
using System.Data;
using System.IO;
using ExcelDataReader;
using OfficeOpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace Sharepoint.Import.Console
{
    internal class Program
    {

        //{"$schema":"https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json","elmType":"div","children":[{"elmType":"a","txtContent":"=if([Category Name]!='', '[Category Name]'+' Documents', '')","attributes":{"target":"_blank","href":"=if([Category Name]!='', 'https://jellinbah.sharepoint.com/SHMS/Shared%20Documents/Forms/AllItems.aspx?id=%2FSHMS%2FShared%20Documents%2F'+'[Category Name]', '')"}}]}
        static void Main(string[] args)
        {
            string siteUrl = "https://jellinbah.sharepoint.com";
            string listName = "Customers";
            string excelFilePath = @"C:\Import-Excel-To-Sharepoint-main\Sharepoint.Import.Console\ExcelToSharePointList\CustomersList.xlsx";

            // Initialize the SharePoint Client Context
            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                // Provide SharePoint credentials
                clientContext.Credentials = new SharePointOnlineCredentials("S2C.Services@jellinbah.com.au", GetSecureString("Sc4*******"));

                // Get the SharePoint list
                List list = clientContext.Web.Lists.GetByTitle(listName);

                // Load the Excel data into a DataTable
                DataTable excelData = ReadExcelWithHeaders(excelFilePath);

                // Loop through the DataTable and add data to the SharePoint list
                foreach (DataRow row in excelData.Rows)
                {
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    ListItem newItem = list.AddItem(itemCreateInfo);

                    // Set column values from the Excel data
                    newItem["Title"] = row["Company Title"];
                    newItem["Market Region"] = row["Market Region"];
                    newItem["Company Location"] = row["Location/Addresses"];
                    newItem["Steel Production"] = row["Steel / Energy Production"];
                    //newItem["Operations"] = row["Operations"];
                    //newItem["Products Purchased"] = row["Products Purchased"];//Manipulate data
                    newItem["Contacts"] = row["Company Contacts"];


                    // Update item
                    newItem.Update();
                }

                // Execute the batch
                clientContext.ExecuteQuery();

                System.Console.WriteLine("Data imported successfully to SharePoint list.");
            }


        }

        // Method to convert plain text password to SecureString
        private static System.Security.SecureString GetSecureString(string password)
        {
            System.Security.SecureString securePassword = new System.Security.SecureString();
            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }
            return securePassword;
        }

        // Method to read Excel data into a DataTable
        //private static DataTable GetExcelData(string filePath)
        //{
        //    //DataTable dataTable = new DataTable();
        //    //using (FileStream stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
        //    //{
        //    //    using (IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
        //    //    {
        //    //        DataSet dataSet = reader.AsDataSet();
        //    //        dataTable = dataSet.Tables[0];
        //    //    }
        //    //}
        //    //return dataTable;
        //    // Load Excel file using EPPlus
        //    FileInfo file = new FileInfo(filePath);
        //    using (ExcelPackage package = new ExcelPackage(file))
        //    {
        //        ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first worksheet

        //        int rowCount = worksheet.Dimension.Rows;
        //        int colCount = worksheet.Dimension.Columns;

        //        // Read headers from the first row
        //        string[] headers = new string[colCount];
        //        for (int col = 1; col <= colCount; col++)
        //        {
        //            headers[col - 1] = worksheet.Cells[1, col].Value?.ToString();
        //        }

        //        // Display headers
        //        System.Console.WriteLine("Headers:");
        //        foreach (string header in headers)
        //        {
        //            System.Console.WriteLine(header);
        //        }

        //        // Read data starting from the second row
        //        System.Console.WriteLine("\nData:");
        //        for (int row = 2; row <= rowCount; row++)
        //        {
        //            for (int col = 1; col <= colCount; col++)
        //            {
        //                string cellValue = worksheet.Cells[row, col].Value?.ToString();
        //                System.Console.Write(cellValue + "\t");
        //            }
        //            System.Console.WriteLine();
        //        }
        //    }
        //}
        static DataTable ReadExcelToDataTable(string filePath)
        {
            DataTable dataTable = new DataTable();

            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first worksheet

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                // Read headers from the first row and add them to the DataTable
                for (int col = 1; col <= colCount; col++)
                {
                    string header = worksheet.Cells[1, col].Value?.ToString();
                    dataTable.Columns.Add(header);
                }

                // Read data starting from the second row and populate the DataTable
                for (int row = 2; row <= rowCount; row++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= colCount; col++)
                    {
                        dataRow[col - 1] = worksheet.Cells[row, col].Value?.ToString();
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }

            return dataTable;
        }

        static DataTable ReadExcelWithHeaders(string filePath)
        {
            DataTable dataTable = new DataTable();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                bool isFirstRow = true;

                foreach (Row row in sheetData.Elements<Row>())
                {
                    if (isFirstRow)
                    {
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            dataTable.Columns.Add(GetValueFromCell(cell, workbookPart));
                        }
                        isFirstRow = false;
                    }
                    else
                    {
                        DataRow dataRow = dataTable.NewRow();
                        int columnIndex = 0;
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            dataRow[columnIndex] = GetValueFromCell(cell, workbookPart);
                            columnIndex++;
                        }
                        dataTable.Rows.Add(dataRow);
                    }
                }
            }

            return dataTable;
        }

        static string GetValueFromCell(Cell cell, WorkbookPart workbookPart)
        {
            string value = cell.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;
                if (stringTablePart != null)
                {
                    value = stringTablePart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }
            return value;
        }


    }
}
