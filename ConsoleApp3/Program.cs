using System;
using System.Xml;
using Aspose.Cells;
using ConsoleApp3;
using ConsoleApp3.Cells;
using ConsoleApp3.ExcelReader;
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

            sheet.PrintDisplay();
            
            
            
            Row? personByRowId = sheet.SearchByRowId(1);
            List<Row> peopleWtihName = sheet.SearchByFirstName("Fallon");
            List<Row> peopleWithNameContaining = sheet.SearchByFirstNameContains("Jo");
            List<Row> peopleWithLastName = sheet.SearchByLastName("Hail");
            List<Row> peopleWithGender = sheet.SearchByGender("Male");
            List<Row> peopleFrom = sheet.SearchByCountry("USA");
            List<Row> peopleWithAge = sheet.SearchByAge(32);
            List<Row> peopleInAgeRange = sheet.SearchByAgeRange(25, 35);
            List<Row> peopleByDate = sheet.SearchByDate("15/10/2017");
            Row? personById = sheet.SearchById(1562);
            List<Row> invalidRows = sheet.SearchOnlyInvalid();
            List<Row> validRows = sheet.SearchOnlyValid();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
        
    } 
}

