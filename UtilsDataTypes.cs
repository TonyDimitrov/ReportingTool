using System;

namespace ReportExtraction
{
    public static class UtilsDataTypes
    {
        private static readonly LoggerConfiguration logger = new LoggerConfiguration();

        private static string logText = "Could not parse db object to {0} value: [{1}]";
        public static int ParseToInt(object value)
        {
            if (int.TryParse(value.ToString(), out var number))
            {
                return number;
            }

            var ex = new ArgumentException(string.Format(logText, "int", value));
            logger.LogError(ex, string.Format(logText, "int", value), false);

            throw ex;
        }

        public static int? ParseToDbNullableInt(object value)
        {
            if (value == DBNull.Value)
            {
                return null;
            }
            if (int.TryParse(value.ToString(), out var number))
            {
                return number;
            }

            return null;
        }

        public static double ParseToDouble(object value)
        {
            if (double.TryParse(value.ToString(), out var number))
            {
                return number;
            }

            logger.LogWarning(string.Format(logText, "double", value), false);

            return 0.0;
        }

        public static float ParseToFloat(object value)
        {
            if (float.TryParse(value.ToString(), out var number))
            {
                return number;
            }

            logger.LogWarning(string.Format(logText, "float", value), false);

            return 0.0f;
        }

        public static float? ParseToDbNullableFloat(object value)
        {
            if (value == DBNull.Value)
            {
                return null;
            }

            if (float.TryParse(value.ToString(), out var number))
            {
                return number;
            }

            logger.LogWarning(string.Format(logText, "float", value), false);

            return null;
        }
    }
}
