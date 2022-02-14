using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentXls.Fluent.Exceptions;

namespace FluentXls.Fluent.Extensions
{
    public static class OpenXmlExtensions
    {
        public static void SetColumnWidth(this Column column, uint columnWidth, uint index)
        {
            column.Width = columnWidth;
            column.Min = new UInt32Value(index);
            column.Max = new UInt32Value(index);
            column.CustomWidth = true;
        }

        public static void SetCellValue(this Cell cell, CellValues dataType, string? value)
        {
            cell.DataType = dataType;
            cell.CellValue = new CellValue(value);
        }

        public static void FillCellFormula(this Cell cell, string formula, int rowNumber, params int[] columnIndexes)
        {
            var coordinates = Array.ConvertAll(columnIndexes, x => CellPositionHelper.ToCoordinates(x) + rowNumber);
            try
            {
                // ReSharper disable once CoVariantArrayConversion
                cell.CellFormula = new CellFormula(string.Format(formula, coordinates));
            }
            catch (FormatException ex)
            {
                throw new MissingFormulaParameterException("Недостаточно параметров для заполнения формулы", ex);
            }
        }
    }
}
