namespace ConsoleApp3.Handlers;
using ConsoleApp3.Sheet;
public class SearchHandler
{
    private readonly Sheet _sheet;

    public SearchHandler(Sheet sheet)
    {
        _sheet = sheet;
    }
    public void StartSearch()
    {
        ConsoleKeyInfo keyinfo;
        string searchWord = "";
        do
        {
            keyinfo = Console.ReadKey(intercept: true);

            if (keyinfo.Key == ConsoleKey.Backspace)
            {
                if (searchWord.Length > 0)
                {
                    searchWord = searchWord.Substring(0, searchWord.Length - 1);
                    
                }
                else
                {
                    continue;
                }
            }
            else
            {
                searchWord += keyinfo.KeyChar;
            }
        
            Console.Clear();
            Console.WriteLine(searchWord);
            _sheet.PrintDisplay(_sheet.SearchByAnyMatch(searchWord));
            
                
        } while (keyinfo.Key != ConsoleKey.Escape);
    }
    
    
}