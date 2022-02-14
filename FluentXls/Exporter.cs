using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentXls.Fluent;
using FluentXls.Fluent.Configuration;
using FluentXls.Fluent.Extensions;

namespace FluentXls
{
    public sealed class Exporter : IDisposable
    {
        private readonly bool _applyHeader;
        private readonly SpreadsheetDocument _document;
        private int _excelId;

        private Exporter(SpreadsheetDocument document, bool applyHeader)
        {
            _document = document;
            _applyHeader = applyHeader;
        }

        public void Dispose() => _document.Dispose();

        private UInt32Value GetNextUIntId()
        {
            Interlocked.Increment(ref _excelId);
            return new UInt32Value((uint)_excelId);
        }

        public Exporter ExportItems<TItem>(
            IEnumerable<TItem> data,
            OpenXmlEntityConfiguration<TItem> columnConfigurations,
            string sheetName,
            bool? applyHeader = null)
        {
            AddSheetWithData(_document.WorkbookPart, GetNextUIntId(), sheetName, data,
                columnConfigurations.GetConfigurations(), applyHeader ?? _applyHeader);
            return this;
        }

        public Exporter ExportItems<TItem>(
            IEnumerable<TItem> data,
            string sheetName,
            bool? applyHeader = null)
        {
            AddSheetWithData(_document.WorkbookPart, GetNextUIntId(), sheetName, data,
                OpenXmlEntityConfiguration.CreateConfigurations(typeof(TItem)), applyHeader ?? _applyHeader);
            return this;
        }

        public Exporter ExportItems<TItem>(
            IEnumerable<TItem> data,
            ICollection<IColumnConfiguration> columnConfigurations,
            string sheetName,
            bool? applyHeader = null)
        {
            AddSheetWithData(_document.WorkbookPart, GetNextUIntId(), sheetName, data, columnConfigurations,
                applyHeader ?? _applyHeader);
            return this;
        }

        public void Build()
        {
            _document.Save();
            _document.Close();
        }

        public static void ExportItems<TItem>(
            Stream stream,
            IEnumerable<TItem> data,
            string sheetName,
            bool applyHeader = true
        ) => ExportItems(stream, data, OpenXmlEntityConfiguration.CreateConfigurations(typeof(TItem)),
            sheetName);

        public static void ExportItems<TItem>(
            Stream stream,
            IEnumerable<TItem> data,
            OpenXmlEntityConfiguration<TItem> columnConfigurations,
            string sheetName,
            bool applyHeader = true
        ) => ExportItems(stream, data, columnConfigurations.GetConfigurations(), sheetName, applyHeader);

        public static void ExportItems<TItem>(
            Stream stream,
            IEnumerable<TItem> data,
            ICollection<IColumnConfiguration> columnConfigurations,
            string sheetName,
            bool applyHeader = true
        )
        {
            using var document = CreateSpreadsheetDocument(stream);
            var exporter = new Exporter(document, applyHeader);
            exporter.ExportItems(data, columnConfigurations, sheetName, applyHeader).Build();
        }

        public static Exporter ExportMulty(
            Stream stream,
            bool applyHeader = true
        )
        {
            var document = CreateSpreadsheetDocument(stream);
            return new Exporter(document, applyHeader);
        }

        private static SpreadsheetDocument CreateSpreadsheetDocument(Stream stream)
        {
            var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            workbookPart.Workbook.Append(new Sheets());
            return document;
        }

        private static void AddSheetWithData<T>(WorkbookPart workbookPart, UInt32Value sheetId, string sheetName,
            IEnumerable<T> data, ICollection<IColumnConfiguration> columnConfigurations, bool applyHeader)
        {
            var strSheetId = $"rId{sheetId}";
            var sheet = new Sheet { Name = sheetName, SheetId = sheetId, Id = strSheetId };

            workbookPart.Workbook.Sheets.Append(sheet);

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>(sheet.Id);
            worksheetPart.Worksheet = new Worksheet();
            CreateColumns(columnConfigurations, worksheetPart.Worksheet);

            var sheetData = new SheetData();
            worksheetPart.Worksheet.Append(sheetData);
            FillSheetData(data, sheetData, columnConfigurations, applyHeader);
        }

        private static void FillSheetData<TItem>(
            IEnumerable<TItem> data,
            SheetData sheetData,
            ICollection<IColumnConfiguration> annotations,
            bool applyHeader
        )
        {
            var rowIndex = 1;
            if (applyHeader)
            {
                var header = FillHeader(annotations);
                sheetData.Append(header);
                rowIndex++;
            }

            foreach (var item in data)
            {
                var row = new Row();
                sheetData.Append(row);
                FillDataRow(item, row, rowIndex++, annotations);
            }
        }

        private static Row FillHeader(ICollection<IColumnConfiguration> annotations)
        {
            var tRow = new Row();
            foreach (var x in annotations)
            {
                var cell = new Cell
                    { CellReference = CellPositionHelper.ToCoordinates(x.ColumnIndex) + 1 };
                tRow.Append(cell);
                cell.SetCellValue(CellValues.String, x.ColumnTitle);
            }

            return tRow;
        }

        private static void FillDataRow<TItem>(TItem item, Row row, int rowIndex,
            ICollection<IColumnConfiguration> annotations)
        {
            foreach (var x in annotations)
            {
                var cell = new Cell
                    { CellReference = CellPositionHelper.ToCoordinates(x.ColumnIndex) + rowIndex };
                row.Append(cell);
                x.FillCell(cell, item!, rowIndex);
            }
        }

        private static void CreateColumns(ICollection<IColumnConfiguration> annotation, Worksheet worksheet)
        {
            var max = annotation.Max(x => x.ColumnIndex);
            var columns = new Columns();
            for (var index = 0; index <= max; index++)
            {
                var columnAnnotation = annotation.FirstOrDefault(x => x.ColumnIndex == index);

                var column = new Column();
                columns.Append(column);
                column.SetColumnWidth(columnAnnotation?.ColumnWidth ?? 10, (uint)(index + 1));
            }

            worksheet.Append(columns);
        }
    }
}
