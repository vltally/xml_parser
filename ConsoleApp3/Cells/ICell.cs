namespace ConsoleApp3.ExcelReader;

public interface ICell //<T>
{
    bool IsValid { get; set; }

    string GetDisplayValue();
    
}