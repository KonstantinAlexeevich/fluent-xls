using System.Linq.Expressions;
using System.Reflection;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentXls.Fluent.Configuration;
using FluentXls.Fluent.ConfigurationBuilders;

namespace FluentXls.Fluent
{
    public class OpenXmlEntityConfiguration<T>
    {
        readonly List<IConfigurationBuilder<T, IColumnConfiguration>> _configurationBuilders =
            new List<IConfigurationBuilder<T, IColumnConfiguration>>();

        List<IColumnConfiguration>? _columnConfigurations;


        public PropertyColumnConfigurationBuilder<T, TProperty> HasPropertyColumn<TProperty>()
        {
            var configuration = new PropertyColumnConfigurationBuilder<T, TProperty>(_configurationBuilders.Count + 1);
            _configurationBuilders.Add(configuration);
            return configuration;
        }

        public PropertyColumnConfigurationBuilder<T, TProperty> HasPropertyColumn<TProperty>(
            Expression<Func<T, TProperty>?> propertyExpression) =>
            HasPropertyColumn<TProperty>().WithProperty(propertyExpression);


        public FormulaColumnConfigurationBuilder<T> HasFormulaColumn()
        {
            var configuration = new FormulaColumnConfigurationBuilder<T>(_configurationBuilders.Count + 1);
            _configurationBuilders.Add(configuration);
            return configuration;
        }

        public FormulaColumnConfigurationBuilder<T> HasFormulaColumn(string formula) =>
            HasFormulaColumn().WithFormula(formula);

        public FormulaColumnConfigurationBuilder<T> HasFormulaColumn(Func<T, int, string> formula) =>
            HasFormulaColumn().WithFormula(formula);

        public CallbackColumnConfigurationBuilder<T, TResult> HasCallbackColumn<TResult>()
        {
            var configuration = new CallbackColumnConfigurationBuilder<T, TResult>(_configurationBuilders.Count + 1);
            _configurationBuilders.Add(configuration);
            return configuration;
        }

        public CallbackColumnConfigurationBuilder<T, TResult> HasCallbackColumn<TResult>(Func<T, TResult> callback) =>
            HasCallbackColumn<TResult>().WithCallback(callback);

        public CallbackColumnConfigurationBuilder<T, TResult> HasCallbackColumn<TResult>(Func<TResult> callback) =>
            HasCallbackColumn<TResult>().WithCallback(x => callback());

        public ICollection<IColumnConfiguration> GetConfigurations()
        {
            if (_columnConfigurations != null)
                return _columnConfigurations;

            _columnConfigurations = _configurationBuilders.Select(x => x.Build(this)).ToList();
            return _columnConfigurations;
        }

        public static string ToCoordinates(int value) => CellPositionHelper.ToCoordinates(value);
        public static int ToColumnIndex(string coordinates) => CellPositionHelper.ToColumnIndex(coordinates);
    }

    public static class OpenXmlEntityConfiguration
    {
        public static OpenXmlEntityConfiguration<T> Create<T>(Action<OpenXmlEntityConfiguration<T>> configurationAction)
        {
            var config = new OpenXmlEntityConfiguration<T>();
            configurationAction(config);
            return config;
        }

        public static ICollection<IColumnConfiguration> CreateConfigurations(Type type) => type
            .GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetProperty)
            .Select(CreateConfiguration)
            .ToList();

        static IColumnConfiguration CreateConfiguration(PropertyInfo propertyInfo, int index) =>
            new CallbackColumnConfiguration(
                index,
                propertyInfo.Name,
                10,
                CellValues.String,
                x => propertyInfo.GetValue(x)?.ToString(),
                propertyInfo.Name
            );
    }
}
