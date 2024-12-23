using System;
using System.Xml;
using Aspose.Cells;
using ConsoleApp3;
using ConsoleApp3.ExcelReader;

class Program
{
    public static void Main()
    {
        
        try
        {
            Workbook workbook = new Workbook("file_example_XLSX_5000.xlsx");
            workbook.Save("output.xml");
            
            
            ExcelReader reader = new ExcelReader();
            List<ExcelRow> data = reader.ReadExcelXml("output.xml");

            Console.WriteLine($"Read {data.Count} rows:");
            foreach (ExcelRow row in data)
            {
                Console.WriteLine($"Row #{row.RowNumber}:");
                Console.WriteLine($"Valid: {row.IsValid}");
                if (!row.IsValid)
                {
                    Console.WriteLine($"Validation issues: {row.ValidationMessage}");
                }
                Console.WriteLine($"First Name: {row.FirstName ?? "not specified"}");
                Console.WriteLine($"Last Name: {row.LastName ?? "not specified"}");
                Console.WriteLine($"Gender: {row.Gender ?? "not specified"}");
                Console.WriteLine($"Country: {row.Country ?? "not specified"}");
                Console.WriteLine($"Age: {row.Age?.ToString() ?? "not specified"}");
                Console.WriteLine($"Date: {row.Date ?? "not specified"}");
                Console.WriteLine($"ID: {row.Id?.ToString() ?? "not specified"}");
                Console.WriteLine("-------------------");
            }

            int validRows = data.Count(r => r.IsValid);
            int invalidRows = data.Count(r => !r.IsValid);
            Console.WriteLine($"\nStatistics:");
            Console.WriteLine($"Valid rows: {validRows}");
            Console.WriteLine($"Invalid rows: {invalidRows}");
            
            Console.WriteLine("-------------------");
            
            List<ExcelRow> invalidRowsList = data.Where(row => !row.IsValid).ToList();
            invalidRowsList.ForEach(row =>
            {
                Console.WriteLine($"Row #{row.RowNumber}:");
                Console.WriteLine($"Valid: {row.IsValid}");
                   Console.WriteLine($"Validation issues: {row.ValidationMessage}");
                Console.WriteLine($"First Name: {row.FirstName ?? "not specified"}");
                Console.WriteLine($"Last Name: {row.LastName ?? "not specified"}");
                Console.WriteLine($"Gender: {row.Gender ?? "not specified"}");
                Console.WriteLine($"Country: {row.Country ?? "not specified"}");
                Console.WriteLine($"Age: {row.Age?.ToString() ?? "not specified"}");
                Console.WriteLine($"Date: {row.Date ?? "not specified"}");
                Console.WriteLine($"ID: {row.Id?.ToString() ?? "not specified"}");
                Console.WriteLine("-------------------");
            });
            
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
        
    } 
}

