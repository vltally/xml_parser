namespace ConsoleApp3.Cells;

public interface ICell //<T>
{
    bool IsValid { get; set; }

    string GetDisplayValue();

    int ColumnIndex {get; }

}