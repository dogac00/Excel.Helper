namespace Excel.Helper.Tests
{
    public class ExcelFilesUtil
    {
        public static string GetExcelsFolderPath(string fileName)
        {
            return $"../../../Excels/{fileName}";
        }
    }
}