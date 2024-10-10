namespace ExcelReport.Driver.Syncfusion
{
    public class WorkbookLoader : IWorkbookLoader
    {
        public IWorkbook Load(string filePath)
        {
            return new Workbook(filePath);
        }
    }
}
