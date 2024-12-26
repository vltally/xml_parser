using System.Xml;
using ConsoleApp3.Cells;
using ConsoleApp3.Rows;

namespace ConsoleApp3.ExcelReader;

public class ExcelReader
{
    private readonly XmlManagerCreator _xmlManagerCreator;

    public ExcelReader(XmlManagerCreator xmlManagerCreator)
    {
        _xmlManagerCreator = xmlManagerCreator;
    }

    public Sheet.Sheet ReadExcelXml(string filePath)
    {
        Sheet.Sheet sheet = new Sheet.Sheet();
        XmlDocument doc = new XmlDocument();
        doc.Load(filePath);

        XmlNamespaceManager nsmgr = _xmlManagerCreator.CreateXmlNamespaceManager(doc.NameTable);
        XmlNodeList? rowNodes = doc.SelectNodes("//def:Worksheet/def:Table/def:Row[position()>1]", nsmgr);

        if (rowNodes == null) return sheet;

        int currentRowIndex = 2;
        
        foreach (XmlNode rowNode in rowNodes)
        {
            Row row = ProcessRow(rowNode, nsmgr, currentRowIndex);
            sheet.Rows.Add(row);
            currentRowIndex++;
        }

        return sheet;
    }

    private Row ProcessRow(XmlNode rowNode, XmlNamespaceManager nsmgr, int rowIndex)
    {
        Row row = new() { RowNumber = rowIndex - 1 };
        Dictionary<int, string> cellValues = ExtractCellValues(rowNode, nsmgr);

        // Створюємо всі 8 клітинок по порядку
        for (int i = 1; i <= 8; i++)
        {
            ICell cell = CreateCell(i, cellValues.GetValueOrDefault(i, string.Empty));
            row.AddCell(cell);
        }

        return row;
    }

    private ICell CreateCell(int columnIndex, string value)
    {
        // RowId, Age та Id - числові значення
        if (columnIndex is 1 or 6 or 8)
        {
            NumberCell cell = new();
            cell.IsValid = int.TryParse(value, out int numValue);
            if (cell.IsValid)
            {
                cell.Value = numValue;
            }
            cell.ColumnIndex = columnIndex;
            return cell;
        }
        
        StringCell stringCell = new()
        {
            Value = value,
            IsValid = !string.IsNullOrEmpty(value),
            ColumnIndex = columnIndex
        };
        return stringCell;
    }
    private Dictionary<int, string> ExtractCellValues(XmlNode rowNode, XmlNamespaceManager nsmgr)
    {
        Dictionary<int, string> cellValues = new Dictionary<int, string>();
        int currentCellIndex = 1;

        XmlNodeList? cellNodes = rowNode.SelectNodes(".//def:Cell", nsmgr);
        if (cellNodes == null) return cellValues;

        foreach (XmlNode cellNode in cellNodes)
        {
            currentCellIndex = GetCellIndex(cellNode, currentCellIndex);
            string? value = ExtractCellValue(cellNode, nsmgr);
            
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

    private string? ExtractCellValue(XmlNode cellNode, XmlNamespaceManager nsmgr)
    {
        XmlNode? dataNode = cellNode.SelectSingleNode(".//def:Data", nsmgr);
        return dataNode?.InnerText;
    }
    
    
}