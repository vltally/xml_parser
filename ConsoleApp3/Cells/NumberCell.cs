using ConsoleApp3.ExcelReader;

namespace ConsoleApp3.Cells;

public class NumberCell : ICell
{
    public int? Value { get; set; }
    
    public bool IsValid { get; set; }
    
    public int ColumnIndex {get; internal set;}
    
    public string GetDisplayValue() => Value.ToString();
}