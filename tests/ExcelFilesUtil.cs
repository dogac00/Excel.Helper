namespace Excel.Helper.Tests
{
    public class ExcelFilesUtil
    {
        public static string GetExcelFilePath(string fileName)
        {
            return $"../../../Excels/{fileName}";
        }
    }
}