using System;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace ExcelWriter
{
	public class ExcelWrapper
	{
		public void CreateFile()
		{
			object missing = System.Reflection.Missing.Value;
			string fileName = $"{DateTime.Now.ToString("ddMMyyyhhmmsstt")}.xlsx";

			string[] columns = { "Default", "Header", "Text", "Is", "Not", "Good", "Enough", "But", "You", "Can", "Write", "Custom", "Column", "Header", "Value" };

			Application excel = new Application();

			if (excel == null)
			{
				Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
				return;
			}
			excel.Visible = false;

			Workbook wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
			Worksheet ws = wb.Worksheets[1];
			//Name the Workbook you are working.
			ws.Name = "First Workbook";

			if (ws == null)
			{
				Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
			}

			Range rng = ws.get_Range("A1", "O1");
			object[] intArray = new object[] { 1, 2, 3, "Some Text", 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15 };
			rng.Value = intArray;
			ListObject table = ws.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, rng, XlYesNoGuess.xlNo);

			table.Name = "MyTableStyle";
			table.TableStyle = "TableStyleLight9";
			for (int i = 1; i <= ws.UsedRange.Columns.Count; i++)
			{
				if (ws.Cells[1, i].Value2 != null)
				{
					ws.Cells[1, i].Value2 = columns[i - 1];
				}
			}

			rng = ws.get_Range("D2", missing);
			Hyperlink hp = ws.Hyperlinks.Add(rng, "https://duckduckgo.com/", string.Empty, "Best Search Engine which doesn't Track you.", "Search Engine");

			try
			{
				try
				{
					//Throw exception if file already exists and user doesn't overwrite the file.
					excel.Application.ActiveWorkbook.SaveAs(fileName);
				
					//Excel shows the file modification dialog to finally Save or Discard the chagnes.
					wb.Close();

					excel.Application.Quit();
					//Throws Exception if Excel window already open with Dialog.
					excel.Quit();
				}
				catch (Exception) { }

				GC.Collect();
				GC.WaitForPendingFinalizers();
				Marshal.FinalReleaseComObject(rng);
				Marshal.FinalReleaseComObject(ws);
				Marshal.FinalReleaseComObject(wb);
				Marshal.FinalReleaseComObject(excel);
			}
			catch (Exception) { }
		}
	}
}