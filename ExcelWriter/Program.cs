namespace ExcelWriter
{
	class Program
	{
		static void Main(string[] args)
		{
			ExcelWrapper ew = new ExcelWrapper();
			ew.CreateFile();
		}
	}
}
