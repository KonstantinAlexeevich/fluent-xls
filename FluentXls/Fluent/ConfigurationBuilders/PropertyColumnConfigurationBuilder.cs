using System.Linq.Expressions;
using FluentXls.Fluent.Configuration;

namespace FluentXls.Fluent.ConfigurationBuilders
{
    public sealed class PropertyColumnConfigurationBuilder<T, TProperty> :
        ConfigurationBuilder<T, TProperty, PropertyColumnConfigurationBuilder<T, TProperty>>
    {
        public PropertyColumnConfigurationBuilder(int columnIndex) => ColumnIndex = columnIndex;

        public PropertyColumnConfigurationBuilder<T, TProperty> WithProperty(
            Expression<Func<T, TProperty>?> propertyExpression)
        {
            WithAction(propertyExpression.Compile()!);
            var propertyName = (propertyExpression.Body as MemberExpression)?.Member.Name;
            WithColumnTitle(propertyName);
            WithColumnKey(propertyName);
            return this;
        }

        public override IColumnConfiguration Build(OpenXmlEntityConfiguration<T> entityConfiguration) =>
            new CallbackColumnConfiguration(
                ColumnIndex,
                ColumnKey,
                (uint)ColumnWidth,
                ColumnType,
                x => Format(Callback!((T)x!)),
                ColumnTitle
            );
    }
}
