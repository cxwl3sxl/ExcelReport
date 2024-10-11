using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Syncfusion.XlsIO;

namespace ExcelReport.Driver.Syncfusion
{
    public class Sheet : ISheet
    {
        private readonly IWorksheet _workbookWorksheet;

        public Sheet(IWorksheet workbookWorksheet)
        {
            _workbookWorksheet = workbookWorksheet;
        }

        public IEnumerator<IRow> GetEnumerator()
        {
            return _workbookWorksheet
                .Rows
                .Select(r => new Row(r))
                .Cast<IRow>()
                .GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public object GetOriginal()
        {
            return _workbookWorksheet;
        }

        public string SheetName => _workbookWorksheet.Name;

        public IRow this[int rowIndex] =>
            rowIndex > _workbookWorksheet.Rows.Length
                ? null
                : new Row(_workbookWorksheet.Rows[rowIndex]);

        public int CopyRows(int start, int end)
        {
            var rowCount = end - start + 1;
            _workbookWorksheet.InsertRow(start + rowCount + 1, rowCount);
            for (var i = start; i <= end; i++)
            {
                var source = _workbookWorksheet.Rows[i];
                var target = _workbookWorksheet.Rows[i + rowCount];
                source.CopyTo(target, ExcelCopyRangeOptions.All);
                target.RowHeight = source.RowHeight;
            }

            return rowCount;
        }

        public int RemoveRows(int start, int end)
        {
            var count = end - start + 1;
            _workbookWorksheet.DeleteRow(start + 1, count);
            return count;
        }
    }
}