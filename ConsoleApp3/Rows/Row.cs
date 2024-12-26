namespace ConsoleApp3.ExcelReader;

public class Row
{
    public int RowNumber { get; set; }
    private readonly Dictionary<int, ICell> _cells = new();
    
    public IReadOnlyDictionary<int, ICell> Cells => _cells;
    
    
    private static readonly int[] ExpectedColumns = { 1, 2, 3, 4, 5, 6, 7, 8 };
    
    public bool IsValid { get; private set; }
    public string ValidationMessage { get; private set; }
    
    public void AddCell(int index, ICell cell)
    {
        _cells[index] = cell;
        Validate();
    }

    private void Validate()
    {
        List<string> messages = new();
       
        foreach (int expectedColumn in ExpectedColumns)
        {
            if (!_cells.ContainsKey(expectedColumn) || !_cells[expectedColumn].IsValid)
            {
                messages.Add($"Column {expectedColumn} is missing or invalid");
            }
        }

        IsValid = !messages.Any();
        ValidationMessage = string.Join("; ", messages);
    }
    
}