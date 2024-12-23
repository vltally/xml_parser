using System.Xml;

namespace ConsoleApp3.ExcelReader;

public class XmlManagerCreator
{
    public XmlNamespaceManager CreateXmlNamespaceManager(XmlNameTable nameTable)
    {
        XmlNamespaceManager nsmgr = new XmlNamespaceManager(nameTable);
        nsmgr.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet");
        nsmgr.AddNamespace("x", "urn:schemas-microsoft-com:office:excel");
        nsmgr.AddNamespace("o", "urn:schemas-microsoft-com:office:office");
        nsmgr.AddNamespace("def", "urn:schemas-microsoft-com:office:spreadsheet");
        return nsmgr;
    }
}