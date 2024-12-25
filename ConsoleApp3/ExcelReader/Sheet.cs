namespace ConsoleApp3.ExcelReader;

public class Sheet
{
    public List<Row> Rows { get; } = new();
    public bool IsValid => Rows.All(x => x.IsValid);
}