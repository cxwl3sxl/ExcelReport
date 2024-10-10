using Syncfusion.XlsIO;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ISyncIWorkbook = Syncfusion.XlsIO.IWorkbook;

namespace ExcelReport.Driver.Syncfusion
{
    public class Workbook : IWorkbook
    {
        private readonly ISyncIWorkbook _workbook;

        public Workbook(string file)
        {
            var excelEngine = new ExcelEngine();
            var application = excelEngine.Excel;
            var fs = File.OpenRead(file);
            _workbook = application.Workbooks.Open(fs);
            fs.Close();
            fs.Dispose();
        }

        public IEnumerator<ISheet> GetEnumerator()
        {
            return _workbook
                .Worksheets
                .Select(t => new Sheet(t))
                .Cast<ISheet>()
                .GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public object GetOriginal()
        {
            return _workbook;
        }

        public ISheet this[string sheetName]
        {
            get
            {
                var s = _workbook
                    .Worksheets
                    .FirstOrDefault(a => a.Name == sheetName);
                return s == null ? null : new Sheet(s);
            }
        }

        public byte[] SaveToBuffer()
        {
            using (var ms = new MemoryStream())
            {
                _workbook.SaveAs(ms);
                return ms.ToArray();
            }
        }
    }
}