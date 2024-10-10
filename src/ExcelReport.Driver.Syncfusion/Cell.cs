using Syncfusion.XlsIO;

namespace ExcelReport.Driver.Syncfusion
{
    public class Cell : ICell
    {
        private readonly IRange _range;

        public Cell(IRange range)
        {
            _range = range;
        }

        public object GetOriginal()
        {
            return _range;
        }

        public int RowIndex => _range.Row - 1;

        public int ColumnIndex => _range.Column - 1;

        public object Value
        {
            get => _range.Value2;
            set => _range.Value2 = value;
        }
    }
}