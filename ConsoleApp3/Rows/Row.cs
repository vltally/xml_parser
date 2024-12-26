using System.Text;
using ConsoleApp3.Cells;
using ConsoleApp3.ExcelReader;

namespace ConsoleApp3.Rows;

public class Row
{
    private readonly List<ICell> _cells = new();
    public IReadOnlyList<ICell> Cells => _cells;
    public int RowNumber { get; set; }
    public bool IsValid { get; private set; }
    public string ValidationMessage { get; private set; }

    
    public void AddCell(ICell cell)
    {
        _cells.Add(cell);
        Validate();
    }

    public static readonly Dictionary<int, string> validationMessages = new()
    {
        { 1, "Missing value in Row ID" },
        { 2, "Missing value in First Name" },
        { 3, "Missing value in Last Name" },
        { 4, "Missing value in Gender" },
        { 5, "Missing value in Country" },
        { 6, "Missing value in Age" },
        { 7, "Missing value in Date" },
        { 8, "Missing value in ID" }
    };
    
    private void Validate()
    {
        List<string> messages = new();
        
        if (_cells.Count != 8)
        {
            messages.Add($"Expected 8 cells, but got {_cells.Count}");
        }

        for (int i = 0; i < _cells.Count; i++)
        {
            if (!_cells[i].IsValid)
            {
                messages.Add($"Cell at position {i + 1} is invalid: {validationMessages[i + 1]}");
            }
        }

        IsValid = !messages.Any();
        ValidationMessage = string.Join("; ", messages);
    }
    
    public string GetDisplay()
    {
        // Проводимо валідацію перед отриманням відображення
        Validate();

        StringBuilder sb = new();

        // Додаємо дані клітинок до StringBuilder
        foreach (ICell cell in _cells)
        {
            sb.Append(cell.GetDisplayValue()).Append("\t | \t");
        }

        // Якщо є повідомлення валідації, додаємо його
        if (!IsValid)
        {
            sb.AppendLine().Append($"Validation: {ValidationMessage}");
        }

        return sb.ToString().Trim();
    }
    
    
}