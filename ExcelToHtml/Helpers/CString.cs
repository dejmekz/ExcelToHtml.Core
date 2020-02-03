namespace ExcelToHtml.Helpers
{
    public static class CString
    {
        #region String extensions

        public static string SubstringAfter(this string value, string afterString)
        {
            int position = value.LastIndexOf(afterString);

            if (position == -1)
                return string.Empty;

            int adjustedPosition = position + afterString.Length;

            if (adjustedPosition >= value.Length)
                return string.Empty;

            return value.Substring(adjustedPosition);
        }

        public static string SubstringBefore(this string value, string beforeSubstring)
        {
            int position = value.IndexOf(beforeSubstring);

            if (position == -1)
                return string.Empty;

            return value.Substring(0, position);
        }

        #endregion String extensions
    }
}