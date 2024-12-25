namespace ConsoleApp3.ExcelReader;

public class Row
{
    public int RowNumber { get; set; }
    public Dictionary<int, ICell> Cells { get; } = new();
    public bool IsValid => Cells.Values.All(x => x.IsValid);
    
    
    
    
    
}