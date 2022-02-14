using DocumentFormat.OpenXml.Spreadsheet;
using FluentXls.Fluent.Extensions;

namespace FluentXls.Fluent.Configuration
{
    public sealed class CallbackColumnConfiguration : IColumnConfiguration
    {
        public CallbackColumnConfiguration(
            int columnIndex,
            string? columnKey,
            uint columnWidth,
            CellValues dataType,
            Func<object, string?> callBack,
            string? columnTitle
        )
        {
            _dataType = dataType;
            _callBack = callBack;
            ColumnIndex = columnIndex;
            ColumnKey = columnKey;
            ColumnWidth = columnWidth;
            ColumnTitle = columnTitle;
        }

        readonly Func<object, string?> _callBack;
        readonly CellValues _dataType;

        public int ColumnIndex { get; }
        public string? ColumnKey { get; }
        public uint? ColumnWidth { get; set; }
        public bool Hidden { get; set; }
        public string? ColumnTitle { get; }

        public void FillCell(Cell cell, object data, int rowNumber) =>
            cell.SetCellValue(_dataType, _callBack.Invoke(data));
    }
}
