using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Common.ExcelUtil
{
	public class HTMLWriteOperation : IWriteOperation
	{
		public StringBuilder Writer { get; set; }

		void IWriteOperation.StartTagWrite(Bold iThing) { Writer.AppendLine("<B>"); }
		void IWriteOperation.EndTagWrite(Bold iThing) { Writer.AppendLine("</B>"); }

		void IWriteOperation.EndTagWrite(Underline aThing) { Writer.AppendLine("</U>"); }
		void IWriteOperation.StartTagWrite(Underline aThing) { Writer.AppendLine("<U>"); }

		void IWriteOperation.StartTagWrite(Italic aThing) { Writer.AppendFormat("<I>"); }
		void IWriteOperation.EndTagWrite(Italic aThing) { Writer.AppendFormat("</I>"); }

		void IWriteOperation.StartTagWrite(Strike u) { Writer.AppendLine("<del>"); }
		void IWriteOperation.EndTagWrite(Strike aThing) { Writer.AppendLine("</del>"); }

		void IWriteOperation.StartTagWrite(Color u)
		{
			if (u.Rgb != null)
			{
				string colorFormat = @"<span style=""color:#{0};"">";
				string span = string.Format(colorFormat, u.Rgb.Value.Substring(2, 6));
				Writer.AppendLine(span);
			}
		}
		void IWriteOperation.EndTagWrite(Color aThing)
		{
			Writer.AppendLine("</span>");
		}
		void IWriteOperation.StartTagWrite(RunFont u)
		{
			if (u.Val.HasValue)
			{
				string fontFormat = @"<span style=""font-family:'{0}';"">";
				string span = string.Format(fontFormat, u.Val.Value);
				Writer.AppendLine(span);
			}
		}
		void IWriteOperation.EndTagWrite(RunFont aThing)
		{
			Writer.AppendLine("</span>");
		}

		void IWriteOperation.EndTagWrite(FontSize aThing)
		{
			Writer.AppendLine("</span>");
		}
		void IWriteOperation.StartTagWrite(FontSize u)
		{
			string fontSizeFormat = @"<span style=""font-size:{0}px;"">";
			string span = string.Format(fontSizeFormat, u.Val);
			Writer.AppendLine(span);
		}

		void IWriteOperation.StartTagWrite(Font u) { }
		void IWriteOperation.EndTagWrite(Font aThing) { }

		void IWriteOperation.EndTagWrite(FontScheme aThing) { }
		void IWriteOperation.StartTagWrite(FontScheme u) { }

		void IWriteOperation.EndTagWrite(Condense aThing) { }
		void IWriteOperation.StartTagWrite(Condense u) { }

		void IWriteOperation.EndTagWrite(VerticalTextAlignment aThing) { }
		void IWriteOperation.StartTagWrite(VerticalTextAlignment u) { }

		void IWriteOperation.EndTagWrite(Shadow aThing) { }
		void IWriteOperation.StartTagWrite(Shadow u) { }

		void IWriteOperation.StartTagWrite(FontFamily u) { }
		void IWriteOperation.EndTagWrite(FontFamily aThing) { }

		void IWriteOperation.StartTagWrite(Outline u) { }
		void IWriteOperation.EndTagWrite(Outline aThing) { }
	}
}
