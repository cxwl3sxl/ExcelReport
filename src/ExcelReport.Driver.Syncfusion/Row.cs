using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Syncfusion.XlsIO;

namespace ExcelReport.Driver.Syncfusion
{
    public class Row : IRow
    {
        private readonly IRange _range;

        public Row(IRange range)
        {
            _range = range;
        }

        public IEnumerator<ICell> GetEnumerator()
        {
            return _range
                .Cells
                .Select(a => new Cell(a))
                .Cast<ICell>()
                .GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public object GetOriginal()
        {
            return _range;
        }

        public ICell this[int columnIndex] =>
            columnIndex < _range.Cells.Length
                ? new Cell(_range.Cells[columnIndex])
                : null;
    }
}