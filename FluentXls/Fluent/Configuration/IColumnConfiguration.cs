using DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXls.Fluent.Configuration
{
    public interface IColumnConfiguration
    {
        int ColumnIndex { get; }
        string? ColumnKey { get; }
        string? ColumnTitle { get; }
        uint? ColumnWidth { get; set; }
        void FillCell(Cell cell, object data, int rowNumber);
    }
}
