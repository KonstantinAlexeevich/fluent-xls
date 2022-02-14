using FluentXls.Fluent.Configuration;

namespace FluentXls.Fluent.ConfigurationBuilders
{
    internal interface IConfigurationBuilder<TEntity, out TConfiguration> where TConfiguration : IColumnConfiguration
    {
        TConfiguration Build(OpenXmlEntityConfiguration<TEntity> entityConfiguration);
    }
}
