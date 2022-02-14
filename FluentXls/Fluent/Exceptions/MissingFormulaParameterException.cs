namespace FluentXls.Fluent.Exceptions
{
    public class MissingFormulaParameterException: Exception
    {
        public MissingFormulaParameterException(): base("Недостаточно параметров для заполнения формулы")
        { }

        public MissingFormulaParameterException(string message, Exception innerException): base(message, innerException)
        { }

        public MissingFormulaParameterException(string columnKey): base($"Недостаточно параметров для заполнения формулы для столбца {columnKey}")
        { }
    }
}
