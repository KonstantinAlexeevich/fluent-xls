using DocumentFormat.OpenXml.Spreadsheet;
using FluentXls.Fluent.Extensions;

namespace FluentXls.Fluent.Configuration
{
    public sealed class FormulaColumnConfiguration<T> : IColumnConfiguration
    {
        public FormulaColumnConfiguration(
            int columnIndex,
            string? columnKey,
            uint columnWidth,
            Func<T?, int, string> formula,
            string? columnTitle,
            OpenXmlEntityConfiguration<T> entityConfiguration,
            List<string?> bindFormulaColumnCodes
        )
        {
            ColumnTitle = columnTitle;
            _formula = formula;
            _entityConfiguration = entityConfiguration;
            ColumnIndex = columnIndex;
            ColumnKey = columnKey;
            ColumnWidth = columnWidth;
            _bindFormulaColumnCodes = bindFormulaColumnCodes;
        }

        readonly List<string?> _bindFormulaColumnCodes;
        readonly Func<T?, int, string> _formula;
        readonly OpenXmlEntityConfiguration<T> _entityConfiguration;

        public int ColumnIndex { get; }
        public string? ColumnKey { get; }
        public uint? ColumnWidth { get; set; }
        public string? ColumnTitle { get; }

        public void FillCell(Cell cell, object? data, int rowNumber)
        {
            var configuredColumns = _entityConfiguration.GetConfigurations();
            var bindCoordinates = configuredColumns
                .Where(x => _bindFormulaColumnCodes.Contains(x.ColumnKey))
                .Select(x => x.ColumnIndex)
                .ToArray();

            cell.FillCellFormula(_formula!.Invoke((T?)data, rowNumber), rowNumber, bindCoordinates);
        }
    }
}
