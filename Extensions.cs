namespace ReportExtraction
{
    public static class Extensions
    {
        public static string AddDirectoryIdentifier(this string firstPart, string secondPart)
        {
            var lastSlashIndex = firstPart.LastIndexOf('\\');

            return firstPart.Insert(lastSlashIndex + 1, secondPart);
        }
    }
}
