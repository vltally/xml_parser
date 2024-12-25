namespace ConsoleApp3.ExcelReader;

public class StringCell : ICell
{
    public string? Value { get; set; }
    public bool IsValid { get; set; }
    public string GetDisplayValue() => Value;
}