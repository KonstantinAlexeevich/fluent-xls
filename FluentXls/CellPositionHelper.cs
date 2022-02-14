namespace FluentXls
{
    internal static class CellPositionHelper
    {
        const string ColumnChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public static int ToColumnIndex(string coordinates)
        {
            var (index, sum) = (0, 0);
            while (!char.IsNumber(coordinates, index))
                sum = sum * 26 + coordinates[index++] - 64;

            return --sum;
        }

        public static string ToCoordinates(int value)
        {
            int dividend = value;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}
