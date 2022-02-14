using System.Linq.Expressions;
using FluentXls.Fluent.Configuration;

namespace FluentXls.Fluent.ConfigurationBuilders
{
    public sealed class FormulaColumnConfigurationBuilder<T>
        : ConfigurationBuilder<T, string, FormulaColumnConfigurationBuilder<T>>
    {
        public FormulaColumnConfigurationBuilder(int columnIndex) => ColumnIndex = columnIndex;

        readonly List<string?> _formulaColumns = new List<string?>();
        Func<T, int, string>? _formula;

        public FormulaColumnConfigurationBuilder<T> WithFormula(string formula)
        {
            _formula = (x, y) => formula;
            return this;
        }

        public FormulaColumnConfigurationBuilder<T> WithFormula(Func<T, int, string> formula)
        {
            _formula = formula;
            return this;
        }


        public FormulaColumnConfigurationBuilder<T> BindFrom(string columnCode)
        {
            _formulaColumns.Add(columnCode);
            return this;
        }

        public FormulaColumnConfigurationBuilder<T> BindFrom<TProperty>(Expression<Func<T, TProperty>> propertyExpression) =>
            BindFrom((propertyExpression.Body as MemberExpression)?.Member.Name!);

        public override IColumnConfiguration Build(OpenXmlEntityConfiguration<T> entityConfiguration) =>
            new FormulaColumnConfiguration<T>(
                ColumnIndex,
                ColumnKey,
                (uint)ColumnWidth,
                _formula!,
                ColumnTitle,
                entityConfiguration,
                _formulaColumns
            );
    }
}
