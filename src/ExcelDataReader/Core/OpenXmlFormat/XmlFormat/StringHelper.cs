using System.Xml;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

internal static class StringHelper
{       
    private const string ElementT = "t";
    private const string ElementR = "r";

    // https://www.w3.org/TR/REC-xml#NT-S
    private static readonly char[] WhitespaceChars = [' ', '\t', '\n', '\r'];

    public static string ReadStringItem(XmlReader reader, string nsSpreadsheetMl)
    {
        if (!XmlReaderHelper.ReadFirstContent(reader))
        {
            return string.Empty;
        }
        
        TextAccumulator text = default;
        while (!reader.EOF)
        {
            if (reader.IsStartElement(ElementT, nsSpreadsheetMl))
            {
                // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                text.Append(ReadElementContent(reader));
            }
            else if (reader.IsStartElement(ElementR, nsSpreadsheetMl))
            {
                ReadRichTextRun(reader, ref text, nsSpreadsheetMl);
            }
            else if (!XmlReaderHelper.SkipContent(reader))
            {
                break;
            }
        }

        return text.ToString();
    }

    private static void ReadRichTextRun(XmlReader reader, ref TextAccumulator text, string nsSpreadsheetMl)
    {
        if (!XmlReaderHelper.ReadFirstContent(reader))
        {
            return;
        }

        while (!reader.EOF)
        {
            if (reader.IsStartElement(ElementT, nsSpreadsheetMl))
            {
                text.Append(ReadElementContent(reader));
            }
            else if (!XmlReaderHelper.SkipContent(reader))
            {
                break;
            }
        }
    }

    private static string ReadElementContent(XmlReader reader)
    {
        if (reader.GetAttribute("xml:space") == "preserve")
            return reader.ReadElementContentAsString();
        else
            return reader.ReadElementContentAsString().Trim(WhitespaceChars);
    }
}
