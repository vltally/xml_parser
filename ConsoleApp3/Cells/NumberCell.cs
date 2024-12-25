namespace ConsoleApp3.ExcelReader;

public class NumberCell : ICell
{
    public int? Value { get; set; }
    
    public bool IsValid { get; set; }

    public string GetDisplayValue() => Value.ToString();
}