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

			object[] tempValues = new object[] { 1, 2, 3, "Some Text", 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15 };

			Range rng = ws.get_Range(GetExcelColumn(1, 1), GetExcelColumn(15, 10));
			int rowsCount = 10;
			int columnsCount = 15;
			object[,] valArray = new object[rowsCount, columnsCount];

			for (int i = 0; i < rowsCount; i++)
			{
				for (int j = 0; j < columnsCount; j++)
				{
					valArray[i, j] = tempValues[j];
				}
			}
			rng.Value = valArray;
			ListObject table = ws.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, rng, XlYesNoGuess.xlNo);
			table.Name = "MyTableStyle";
			table.TableStyle = "TableStyleLight9";

			for (int i = 0; i < rowsCount; i++)
			{
				rng = ws.get_Range("D" + (2 + i), missing);
				Hyperlink hp = ws.Hyperlinks.Add(rng, "https://duckduckgo.com/", string.Empty, "Best Search Engine which doesn't Track you.", "Search Engine");
			}

			for (int i = 1; i <= ws.UsedRange.Columns.Count; i++)
			{
				if (ws.Cells[1, i].Value2 != null)
				{
					ws.Cells[1, i].Value2 = columns[i - 1];
				}
			}

			try
			{
				try
				{
					//Catch block for exception if file already exists and user doesn't overwrite the file.
					//Default location to save file in 'Documents' folder.
					//You can write absolute path to save the file.
					// string fileNameWithAbsolutePath = "C:\some-folder\" + fileName;
					excel.Application.ActiveWorkbook.SaveAs(fileName);

					//Catch block for exception for if user doesn't save the file when Excel shows the file modification dialog to Save or Discard the chagnes.
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

		private string GetExcelColumn(int columnNumber, int rowNumber)
		{
			return GetExcelColumnString(columnNumber) + rowNumber;
		}

		private string GetExcelColumnString(int columnNumber)
		{
			int dividend = columnNumber;
			string columnName = String.Empty;
			int modulo;

			while (dividend > 0)
			{
				modulo = (dividend - 1) % 26;
				columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
				dividend = (int)((dividend - modulo) / 26);
			}

			return columnName;
		}
	}
}