using DocumentFormat.OpenXml.Spreadsheet;

namespace Common.ExcelUtil
{
	public interface IWriteOperation
	{
		void StartTagWrite(Bold bold);
		void StartTagWrite(Italic it);
		void StartTagWrite(Underline underline);
		void StartTagWrite(Color c);
		void StartTagWrite(Font font);
		void StartTagWrite(FontSize fontSize);
		void StartTagWrite(FontFamily fontFamily);
		void StartTagWrite(FontScheme fontScheme);
		void StartTagWrite(RunFont runFont);
		void StartTagWrite(Strike strike);
		void StartTagWrite(VerticalTextAlignment vertialAlignment);
		void StartTagWrite(Shadow shadow);
		void StartTagWrite(Outline outline);
		void StartTagWrite(Condense condense);

		void EndTagWrite(Bold b);
		void EndTagWrite(Italic b);
		void EndTagWrite(Underline b);
		void EndTagWrite(Color u);
		void EndTagWrite(Font u);
		void EndTagWrite(FontSize u);
		void EndTagWrite(FontFamily u);
		void EndTagWrite(RunFont u);
		void EndTagWrite(FontScheme u);
		void EndTagWrite(Strike u);
		void EndTagWrite(VerticalTextAlignment u);
		void EndTagWrite(Shadow u);
		void EndTagWrite(Outline u);
		void EndTagWrite(Condense u);
	}
}
