using System.Xml;

namespace ConsoleApp3.ExcelReader;
public class ExcelReader
{
    private readonly Dictionary<int, Func<string, ExcelRow, bool>> _cellProcessors;

    public ExcelReader()
    {
        _cellProcessors = InitializeCellProcessors();
    }
    
    private Dictionary<int, string> ExtractCellValues(XmlNode rowNode, XmlNamespaceManager nsmngr)
    {
        Dictionary<int, string> cellValues = new Dictionary<int, string>();
        int currentCellIndex = 1;
    
        XmlNodeList? cellNodes = rowNode.SelectNodes(".//def:Cell", nsmngr);
        if (cellNodes == null) return cellValues;
    
        foreach (XmlNode cellNode in cellNodes)
        {
            currentCellIndex = GetCellIndex(cellNode, currentCellIndex);
            string? value = ExtractCellValue(cellNode, nsmngr);
            
            if (value != null)
            {
                cellValues[currentCellIndex] = value;
            }
            
            currentCellIndex++;
        }
    
        return cellValues;
    }
    
    private int GetCellIndex(XmlNode cellNode, int currentIndex)
    {
        XmlAttribute? indexAttr = cellNode.Attributes?["ss:Index"];
        return indexAttr != null ? int.Parse(indexAttr.Value) : currentIndex;
    }

    private string? ExtractCellValue(XmlNode cellNode, XmlNamespaceManager nsmngr)
    {
        XmlNode? dataNode = cellNode.SelectSingleNode(".//def:Data", nsmngr);
        return dataNode?.InnerText;
    }
    
    private void ConfigureNamespaces(XmlNamespaceManager nsmgr)
    {
        Dictionary<string, string> namespaces = new Dictionary<string, string>
        {
            {"ss", "urn:schemas-microsoft-com:office:spreadsheet"},
            {"x", "urn:schemas-microsoft-com:office:excel"},
            {"o", "urn:schemas-microsoft-com:office:office"},
            {"def", "urn:schemas-microsoft-com:office:spreadsheet"}
        };

        foreach (KeyValuePair<string, string> ns in namespaces)
        {
            nsmgr.AddNamespace(ns.Key, ns.Value);
        }
    }
    
    private void ValidateRow(ExcelRow row, List<string> validationMessages)
    {
        Dictionary<string, string?> requiredFields = new Dictionary<string, string?>
        {
            {"First name", row.FirstName},
            {"Last name", row.LastName},
            {"Gender", row.Gender},
            {"Country", row.Country},
            {"Age", row.Age.ToString()}
        };

        foreach (KeyValuePair<string, string?> field in requiredFields)
        {
            if (string.IsNullOrEmpty(field.Value))
            {
                validationMessages.Add($"{field.Key} is missing");
            }
        }

        if (validationMessages.Any())
        {
            row.IsValid = false;
            row.ValidationMessage = string.Join("; ", validationMessages.Distinct());
        }
    }
    
    private void HandleRowProcessingError(ExcelRow row, Exception ex)
    {
        row.IsValid = false;
        row.ValidationMessage = $"Row processing error: {ex.Message}";
    }
    
    private Dictionary<int, Func<string, ExcelRow, bool>> InitializeCellProcessors()
    {
        return new Dictionary<int, Func<string, ExcelRow, bool>>
        {
            {2, (value, row) => { row.FirstName = value; return !string.IsNullOrEmpty(value); }},
            {3, (value, row) => { row.LastName = value; return !string.IsNullOrEmpty(value); }},
            {4, (value, row) => { row.Gender = value; return !string.IsNullOrEmpty(value); }},
            {5, (value, row) => { row.Country = value; return !string.IsNullOrEmpty(value); }},
            {6, (value, row) => { return int.TryParse(value, out int age) && (row.Age = age) >= 0; }},
            {7, (value, row) => { row.Date = value; return !string.IsNullOrEmpty(value); }},
            {8, (value, row) => { return int.TryParse(value, out int id) && (row.Id = id) >= 0; }}

        };
    }

    private void ProcessCellValues(Dictionary<int, string> cellValues, ExcelRow row, int currentRowIndex,
        List<string> validationMessages)
    {
        row.RowNumber = currentRowIndex - 1;

        foreach (KeyValuePair<int, string> cell in cellValues)
        {
            if (_cellProcessors.TryGetValue(cell.Key, out Func<string, ExcelRow, bool>? processor))
            {
                if (!processor(cell.Value, row))
                {
                    validationMessages.Add($"Invalid value in column {cell.Key}: {cell.Value}");
                    row.IsValid = false;
                }
            }
        }
    }

    public List<ExcelRow> ReadExcelXml(string filePath)
    {
        List<ExcelRow> rows = new List<ExcelRow>();
       
        //
        XmlDocument doc = new XmlDocument();
        doc.Load(filePath);

        //
        XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
        ConfigureNamespaces(nsmgr);
        XmlNodeList rowNodes = doc.SelectNodes("//def:Worksheet/def:Table/def:Row[position()>1]", nsmgr);

        
        
        if (rowNodes != null)
        {
            int currentRowIndex = 2; // Start from 2 as 1 is header
            foreach (XmlNode rowNode in rowNodes)
            {
                ExcelRow row = new ExcelRow();
                row.IsValid = true;
                List<string> validationMessages = new List<string>();

                try
                {
                    ProcessCellValues(ExtractCellValues(rowNode, nsmgr), row, currentRowIndex, validationMessages);
                    ValidateRow(row, validationMessages);
                }
                catch (Exception ex)
                {
                    HandleRowProcessingError(row, ex);
                }

                rows.Add(row);
                currentRowIndex++;
            }
        }

        return rows;
    }
}