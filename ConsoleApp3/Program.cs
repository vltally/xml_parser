using System;
using System.Xml;
using Aspose.Cells;
using ConsoleApp3;
using ConsoleApp3.Cells;
using ConsoleApp3.ExcelReader;
using ConsoleApp3.Handlers;
using ConsoleApp3.Sheet;
using Row = ConsoleApp3.Rows.Row;

class Program
{
    public static void Main()
    {
        
        try
        {
            // Convert Excel to XML
            Workbook workbook = new Workbook("file_example_XLSX_5000.xlsx");
            workbook.Save("output.xml");
            
            // Read and process the XML
            XmlManagerCreator xmlManagerCreator = new XmlManagerCreator();
            ExcelReader reader = new ExcelReader(xmlManagerCreator);
            Sheet sheet = reader.ReadExcelXml("output.xml");

           // sheet.PrintDisplay();
            
            SearchHandler searchHandler = new SearchHandler(sheet);
            searchHandler.StartSearch();
            
            
            
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
        
    } 
}

