using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Common.ExcelUtil
{
    public class SharedStringUtil
    {
	    public static object ReadFromCell(Cell cell)
	    {
		    if (cell == null) return null;
		    if (cell.DataType == null)
			    return cell.InnerText;
		    switch (cell.DataType.Value)
		    {
			    case CellValues.SharedString:
				    return ReadSharedString(cell);
				case CellValues.Boolean:
				    switch (cell.InnerText)
				    {
					    case "0":
						    return false;
						case "1":
						    return true;
				    }
				    break;
		    }
		    return cell.InnerText;
	    }

	    private static string ReadSharedString(Cell cell)
	    {
		    string value = cell.InnerText;
		    SharedStringTablePart sstp = GetSharedTablePart(cell);
		    if (sstp != null && sstp.SharedStringTable!=null)
		    {
			    int index;
			    if (int.TryParse(value, out index))
			    {
				    if (sstp.SharedStringTable.ElementAt(index).ChildElements.Count == 0)
				    {
					    value = sstp.SharedStringTable.ElementAt(index).InnerText;
				    }
				    else
				    {
					    // collection of runs 
					    value = GetStringValueFromRuns(sstp.SharedStringTable.ElementAt(index).ChildElements);
				    }
			    }
		    }
		    return value;
	    }

	    private static string GetStringValueFromRuns(OpenXmlElementList openXmlElementList)
	    {
		    StringBuilder sb = new StringBuilder();
		    foreach (Run run in openXmlElementList.OfType<Run>())
		    {
			    if (run.RunProperties != null)
			    {
				    foreach (OpenXmlLeafElement element in run.RunProperties.OrderBy(e => e.LocalName).OfType<OpenXmlLeafElement>())
				    {
					    element.StartTagWrite(sb);
				    }
			    }
			    sb.Append(run.Text.Text.Replace("\n", "<br/>"));
				if (run.RunProperties != null)
				{
					foreach (OpenXmlLeafElement element in run.RunProperties.OrderBy(e => e.LocalName).OfType<OpenXmlLeafElement>())
					{
						element.EndTagWrite(sb);
					}
				}
		    }
		    return sb.ToString();
	    }

	    private static SharedStringTablePart GetSharedTablePart(Cell cell)
	    {
		    Worksheet ws = cell.Ancestors<Worksheet>().FirstOrDefault();
		    if (ws != null)
		    {
			    var doc = ws.WorksheetPart.OpenXmlPackage as SpreadsheetDocument;
			    if (doc != null)
			    {
				    SharedStringTablePart sstp = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
				    return sstp;
			    }
		    }
		    return null;
	    }
    }

	static class Extension
	{
		public static void StartTagWrite(this OpenXmlLeafElement runElement, StringBuilder sb)
		{
			IWriteOperation x = new HTMLWriteOperation {Writer = sb};
			x.StartTagWrite((dynamic)runElement);
		}

		public static void EndTagWrite(this OpenXmlLeafElement runElement, StringBuilder sb)
		{
			IWriteOperation x = new HTMLWriteOperation { Writer = sb };
			x.EndTagWrite((dynamic)runElement);
		}
	}
}
