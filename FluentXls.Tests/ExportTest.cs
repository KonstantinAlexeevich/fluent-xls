using System;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentXls.Fluent;
using Xunit;

namespace FluentXls.Tests
{
    public class ExportTest
    {
        static ExportItem[] CreateItems()
        {
            var item1 = new ExportItem
            {
                Code = "Code Code",
                Date = DateTime.Now,
                Number = 1,
                Decimal = 101.1M
            };

            var item2 = new ExportItem
            {
                Code = "Code1",
                Date = DateTime.Now,
                Number = 2,
                Decimal = 1.01M
            };

            return new[] {item1, item2};
        }

        [Fact]
        public void ExportSheetsManyTest()
        {
            var items = CreateItems();
            using var memoryStream = new MemoryStream();

            Exporter.ExportMulty(memoryStream)
                .ExportItems(items, "Лист 1", true)
                .ExportItems(items, "Лист 2", false)
                .ExportItems(items, "Лист 3", true)
                .Build();

            memoryStream.Seek(0, SeekOrigin.Begin);
            File.WriteAllBytes("Export.xlsx", memoryStream.ToArray());
        }


        [Fact]
        public void ExportSheetsManyTestFluent()
        {
            var items = CreateItems();
            using var memoryStream = new MemoryStream();

            Exporter.ExportMulty(memoryStream)
                .ExportItems(items, new ExportItemConfiguration(), "Лист 1", true)
                .ExportItems(items, new ExportItemConfiguration(), "Лист 2", false)
                .ExportItems(items, new ExportItemConfiguration(), "Лист 3", true)
                .Build();

            memoryStream.Seek(0, SeekOrigin.Begin);
            File.WriteAllBytes("Export.xlsx", memoryStream.ToArray());
        }

        [Fact]
        public void ExportSheetTest()
        {
            var items = CreateItems();
            using var memoryStream = new MemoryStream();
            Exporter.ExportItems(memoryStream, items, "Лист 1");
            memoryStream.Seek(0, SeekOrigin.Begin);
            File.WriteAllBytes("Export.xlsx", memoryStream.ToArray());
        }

        [Fact]
        public void ExportSheetTestFluent()
        {
            var items = CreateItems();
            using var memoryStream = new MemoryStream();
            Exporter.ExportItems(memoryStream, items, new ExportItemConfiguration(), "Лист 1");
            memoryStream.Seek(0, SeekOrigin.Begin);
            File.WriteAllBytes("Export.xlsx", memoryStream.ToArray());
        }
    }

    public class ExportItem
    {
        public string? Code { get; set; }
        public DateTime? Date { get; set; }
        public int Number { get; set; }
        public decimal Decimal { get; set; }
    }

    public class ExportItemConfiguration : OpenXmlEntityConfiguration<ExportItem>
    {
        public ExportItemConfiguration()
        {
            HasPropertyColumn(x => x.Code)
                .WithColumnTitle("Код")
                .WithColumnWidth(40);

            HasFormulaColumn("SUM({0},{1})")
                .WithColumnTitle("Дельта")
                .WithColumnKey("delta")
                .BindFrom(x => x.Number)
                .BindFrom(x => x.Decimal);

            HasPropertyColumn(x => x.Number)
                .WithColumnType(CellValues.Number)
                .WithColumnWidth(20);

            HasPropertyColumn(x => x.Decimal)
                .WithColumnWidth(20)
                .WithFormat(x => x.ToString("0.00", CultureInfo.InvariantCulture))
                .WithColumnType(CellValues.Number);

            HasFormulaColumn("{0} + {1} + 1")
                .WithColumnIndex(7)
                .WithColumnWidth(50)
                .WithColumnTitle("Дельта + Decimal + 1")
                .BindFrom("delta")
                .BindFrom(x => x.Decimal);

            HasCallbackColumn(() => 1)
                .WithColumnIndex(8)
                .WithColumnType(CellValues.Number);

            HasFormulaColumn((x, y) => $"SUM({ToCoordinates(10)}{y}:KH{y})")
                .WithColumnIndex(9)
                .WithColumnWidth(50)
                .WithColumnTitle("Сумма по всей строке")
                .BindFrom("delta");
        }
    }
}
