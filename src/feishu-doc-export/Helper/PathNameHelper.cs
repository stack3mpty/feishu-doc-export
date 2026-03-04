using System.Text.RegularExpressions;

namespace feishu_doc_export.Helper
{
    public static class PathNameHelper
    {
        private static readonly HashSet<string> ReservedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "CON", "PRN", "AUX", "NUL",
            "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
        };

        public static string SanitizePathSegment(string rawName, string fallback = "untitled")
        {
            var name = rawName ?? string.Empty;
            name = Regex.Replace(name, @"[\\/:*?""<>|\x00-\x1F]", "-");
            name = name.Trim();
            name = name.TrimEnd('.', ' ');

            if (string.IsNullOrWhiteSpace(name))
            {
                name = fallback;
            }

            if (ReservedFileNames.Contains(name))
            {
                name = "_" + name;
            }

            return name;
        }
    }
}
