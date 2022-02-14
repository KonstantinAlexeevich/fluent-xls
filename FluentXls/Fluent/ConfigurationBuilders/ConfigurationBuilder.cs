using DocumentFormat.OpenXml.Spreadsheet;
using FluentXls.Fluent.Configuration;

namespace FluentXls.Fluent.ConfigurationBuilders
{
    public abstract class ConfigurationBuilder<T, TResult, TBuilder>: IConfigurationBuilder<T, IColumnConfiguration>
        where TBuilder: ConfigurationBuilder<T, TResult, TBuilder>
    {
        protected Func<T, TResult>? Callback { get; set; }
        protected int ColumnIndex { get; set; }
        protected string? ColumnKey { get; set; } = Guid.NewGuid().ToString();
        protected string? ColumnTitle { get; set; }
        protected CellValues ColumnType { get; set; } = CellValues.String;
        protected int ColumnWidth { get; set; } = 10;
        protected Func<TResult, string?> Format { get; set; } = x => x?.ToString();

        public TBuilder WithColumnTitle(string? title)
        {
            ColumnTitle = title;
            return (TBuilder)this;
        }

        public TBuilder WithColumnWidth(int width)
        {
            ColumnWidth = width;
            return (TBuilder)this;
        }

        public TBuilder WithColumnIndex(int columnIndex)
        {
            ColumnIndex = columnIndex;
            return (TBuilder)this;
        }

        public TBuilder WithColumnKey(string? code)
        {
            ColumnKey = code;
            return (TBuilder)this;
        }

        public TBuilder WithColumnType(CellValues dataType)
        {
            ColumnType = dataType;
            return (TBuilder)this;
        }

        public TBuilder WithFormat(Func<TResult, string> format)
        {
            Format = format;
            return (TBuilder)this;
        }

        protected TBuilder WithAction(Func<T, TResult> callback)
        {
            Callback = callback;
            return (TBuilder)this;
        }

        public abstract IColumnConfiguration Build(OpenXmlEntityConfiguration<T> entityConfiguration);
    }
}
