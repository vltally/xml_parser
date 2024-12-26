using ConsoleApp3.Cells;
using ConsoleApp3.ExcelReader;
using ConsoleApp3.Rows;

namespace ConsoleApp3.Sheet;

public class Sheet
{
    public List<Row> Rows { get; } = new();
    public bool IsValid => Rows.All(x => x.IsValid);
    public Row? SearchByRowId(int rowId)
    {
        return Rows.FirstOrDefault(row => 
            row.Cells[0] is NumberCell numberCell && 
            numberCell.Value == rowId);
    }

   
    public Row? SearchById(int id)
    {
        return Rows.FirstOrDefault(row => 
            row.Cells[7] is NumberCell numberCell && 
            numberCell.Value == id);
    }

    public List<Row> SearchByFirstName(string firstName)
    {
        return Rows.Where(row =>
            row.Cells[1] is StringCell stringCell &&
            stringCell.Value.Equals(firstName, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    
    public List<Row> SearchByLastName(string lastName)
    {
        return Rows.Where(row =>
            row.Cells[2] is StringCell stringCell &&
            stringCell.Value.Equals(lastName, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }


    public List<Row> SearchByGender(string gender)
    {
        return Rows.Where(row =>
            row.Cells[3] is StringCell stringCell &&
            stringCell.Value.Equals(gender, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    
    public List<Row> SearchByCountry(string country)
    {
        return Rows.Where(row =>
            row.Cells[4] is StringCell stringCell &&
            stringCell.Value.Equals(country, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    
    public List<Row> SearchByAge(int age)
    {
        return Rows.Where(row =>
            row.Cells[5] is NumberCell numberCell &&
            numberCell.Value == age)
            .ToList();
    }

   
    public List<Row> SearchByDate(string date)
    {
        return Rows.Where(row =>
            row.Cells[6] is StringCell stringCell &&
            stringCell.Value.Equals(date))
            .ToList();
    }

   
    public List<Row> SearchByAgeRange(int minAge, int maxAge)
    {
        return Rows.Where(row =>
            row.Cells[5] is NumberCell numberCell &&
            numberCell.Value >= minAge &&
            numberCell.Value <= maxAge)
            .ToList();
    }

   
    public List<Row> SearchByFirstNameContains(string namepart)
    {
        return Rows.Where(row =>
            row.Cells[1] is StringCell stringCell &&
            stringCell.Value.Contains(namepart, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    public List<Row> SearchOnlyInvalid()
    {
        return Rows.Where(row => row.IsValid == false).ToList();
    }
    
    public List<Row> SearchOnlyValid()
    {
        return Rows.Where(row => row.IsValid == true).ToList();
    }
    
    public void PrintDisplay()
    {
        foreach (var row in Rows)
        {
            Console.WriteLine($"Row #{row.RowNumber}: {row.GetDisplay()}");
            Console.WriteLine("----------------------------------------------------------------");
        }
        Console.WriteLine("-------------------");
        Console.WriteLine($"Total rows: {Rows.Count}");
        Console.WriteLine($"Valid rows: {Rows.Count(r => r.IsValid)}");
        Console.WriteLine($"Invalid rows: {Rows.Count(r => !r.IsValid)}");
        Console.WriteLine("-------------------");
    }
    
    
}