using FluentXls.Fluent.Configuration;

namespace FluentXls.Fluent.ConfigurationBuilders
{
    public sealed class CallbackColumnConfigurationBuilder<T, TResult>
        : ConfigurationBuilder<T, TResult, CallbackColumnConfigurationBuilder<T, TResult>>
    {
        public CallbackColumnConfigurationBuilder(int columnIndex) => ColumnIndex = columnIndex;

        public CallbackColumnConfigurationBuilder<T, TResult> WithCallback(Func<T, TResult> callback)
            => WithAction(callback);

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
