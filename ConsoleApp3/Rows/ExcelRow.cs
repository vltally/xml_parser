namespace ConsoleApp3.ExcelReader;

public class ExcelRow
{
    public int RowNumber { get; set; }
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public string Gender { get; set; }
    public string Country { get; set; }
    public int? Age { get; set; }
    public string Date { get; set; }
    public int? Id { get; set; }
    public bool IsValid { get; set; }
    public string ValidationMessage { get; set; }
}