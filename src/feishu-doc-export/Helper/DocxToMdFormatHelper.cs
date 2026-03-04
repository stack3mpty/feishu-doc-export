using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace feishu_doc_export.Helper
{
    public static class DocxToMdFormatHelper
    {
        public static string ReplaceImagePath(this string markdownContent, string currentDocPath)
        {
            var regex = new Regex(@"!\[.*?\]\((.*?)\)", RegexOptions.IgnoreCase);

            var replacedContent = regex.Replace(markdownContent, match =>
            {
                var imagePath = match.Groups[1].Value.Trim();
                var normalizedPath = NormalizeLocalImagePath(imagePath);
                if (string.IsNullOrWhiteSpace(normalizedPath))
                {
                    return match.Value;
                }

                if (Path.IsPathRooted(normalizedPath))
                {
                    var relativePath = Path.GetRelativePath(Path.GetDirectoryName(currentDocPath) ?? ".", normalizedPath);
                    return $"![...]({relativePath})";
                }

                return match.Value;
            });

            return replacedContent;
        }

        public static string ReplaceDocRefPath(this string markdownContent, string currentDocPath)
        {
            var regex = new Regex(@"\[(?<linkText>[^\]]+)\]\((https?://[^/\s]+\.(?:feishu\.cn|larksuite\.com)/wiki/(?<nodeToken>[^)\?#/]+)[^)]*)\)", RegexOptions.IgnoreCase);

            var replacedContent = regex.Replace(markdownContent, match =>
            {
                var fileExt = Path.GetExtension(currentDocPath);
                var nodeToken = match.Groups["nodeToken"].Value;
                var linkText = match.Groups["linkText"].Value;

                var refDocPath = DocumentPathGenerator.GetDocumentPathByNodeToken(nodeToken);
                if (!string.IsNullOrWhiteSpace(refDocPath))
                {
                    var relativePath = Path.GetRelativePath(Path.GetDirectoryName(currentDocPath) ?? ".", refDocPath);
                    return $"[{linkText}]({relativePath}{fileExt})";
                }

                return match.Value;
            });

            return replacedContent;
        }

        public static string ReplaceCodeToMdFormat(this string markdownContent)
        {
            const string pattern = @"\|(?<content>[^\n]+)\n\|\s*:\s*-\s*\|";

            var replacedContent = Regex.Replace(markdownContent, pattern, match =>
            {
                string replacement = match.Groups["content"].Value.Replace("<br>", "\n");
                replacement = replacement.Remove(replacement.LastIndexOf('|'), 1);
                replacement = replacement.Replace("`", string.Empty);
                return $"```{replacement}```";
            });

            return replacedContent;
        }

        public static string CleanupExportArtifacts(this string markdownContent, string currentDocPath)
        {
            if (string.IsNullOrWhiteSpace(markdownContent))
            {
                return markdownContent;
            }

            var lines = markdownContent.Replace("\r\n", "\n").Split('\n').ToList();
            var removeFlags = new bool[lines.Count];
            var hasAsposeWatermark = false;

            for (var i = 0; i < lines.Count; i++)
            {
                var trimmed = lines[i].Trim();
                if (!IsAsposeWatermarkLine(trimmed))
                {
                    continue;
                }

                hasAsposeWatermark = true;
                removeFlags[i] = true;

                var prevLineIndex = FindPrevNonEmptyLine(lines, removeFlags, i - 1);
                if (prevLineIndex >= 0 && IsMarkdownImageLine(lines[prevLineIndex].Trim()))
                {
                    removeFlags[prevLineIndex] = true;
                    TryDeleteImageByMarkdownLine(lines[prevLineIndex], currentDocPath);
                }
            }

            RemoveLikelyAsposeLeadingImageLines(lines, removeFlags, currentDocPath, hasAsposeWatermark);
            RemoveLeadingTitleLine(lines, removeFlags, currentDocPath);

            var keptLines = new List<string>();
            var previousIsBlank = false;
            for (var i = 0; i < lines.Count; i++)
            {
                if (removeFlags[i])
                {
                    continue;
                }

                var line = lines[i];
                var currentIsBlank = string.IsNullOrWhiteSpace(line);
                if (currentIsBlank && previousIsBlank)
                {
                    continue;
                }

                keptLines.Add(line);
                previousIsBlank = currentIsBlank;
            }

            while (keptLines.Any() && string.IsNullOrWhiteSpace(keptLines.First()))
            {
                keptLines.RemoveAt(0);
            }

            while (keptLines.Any() && string.IsNullOrWhiteSpace(keptLines.Last()))
            {
                keptLines.RemoveAt(keptLines.Count - 1);
            }

            return string.Join(Environment.NewLine, keptLines);
        }

        private static bool IsAsposeWatermarkLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                return false;
            }

            var normalized = line.ToLowerInvariant();
            if (normalized.Contains("this document was truncated here because it was created in the evaluation mode"))
            {
                return true;
            }

            if (!normalized.Contains("aspose.words"))
            {
                return false;
            }

            return normalized.Contains("evaluation only")
                   || normalized.Contains("created with an evaluation copy")
                   || normalized.Contains("evaluation mode")
                   || normalized.Contains("to discover the full versions of our apis");
        }

        private static int FindPrevNonEmptyLine(List<string> lines, bool[] removeFlags, int startIndex)
        {
            for (var i = startIndex; i >= 0; i--)
            {
                if (removeFlags[i])
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(lines[i]))
                {
                    continue;
                }

                return i;
            }

            return -1;
        }

        private static bool IsMarkdownImageLine(string line)
        {
            return Regex.IsMatch(line, @"^!\[.*?\]\((.*?)\)\s*$", RegexOptions.IgnoreCase);
        }

        private static void TryDeleteImageByMarkdownLine(string markdownImageLine, string currentDocPath)
        {
            try
            {
                var matches = Regex.Matches(markdownImageLine, @"!\[.*?\]\((.*?)\)", RegexOptions.IgnoreCase);
                foreach (Match match in matches)
                {
                    TryDeleteImageByPath(match.Groups[1].Value, currentDocPath);
                }
            }
            catch
            {
            }
        }

        private static void RemoveLikelyAsposeLeadingImageLines(List<string> lines, bool[] removeFlags, string currentDocPath, bool hasAsposeWatermark)
        {
            if (!hasAsposeWatermark || lines.Count == 0)
            {
                return;
            }

            var maxScanLines = Math.Min(lines.Count, 12);
            for (var i = 0; i < maxScanLines; i++)
            {
                if (removeFlags[i])
                {
                    continue;
                }

                var currentLine = lines[i].Trim();
                if (string.IsNullOrWhiteSpace(currentLine))
                {
                    continue;
                }

                if (!TryGetImagePathsFromImageOnlyLine(currentLine, out var imagePaths))
                {
                    break;
                }

                if (!imagePaths.All(path => IsLikelyAsposeLeadingImagePath(path, currentDocPath)))
                {
                    break;
                }

                removeFlags[i] = true;
                foreach (var imagePath in imagePaths)
                {
                    TryDeleteImageByPath(imagePath, currentDocPath);
                }
            }
        }

        private static bool TryGetImagePathsFromImageOnlyLine(string line, out List<string> imagePaths)
        {
            imagePaths = Regex.Matches(line, @"!\[.*?\]\((.*?)\)", RegexOptions.IgnoreCase)
                .Cast<Match>()
                .Select(x => x.Groups[1].Value.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            if (!imagePaths.Any())
            {
                return false;
            }

            var contentWithoutImages = Regex.Replace(line, @"!\[.*?\]\((.*?)\)", string.Empty, RegexOptions.IgnoreCase).Trim();
            return string.IsNullOrWhiteSpace(contentWithoutImages);
        }

        private static bool IsLikelyAsposeLeadingImagePath(string imagePath, string currentDocPath)
        {
            var normalizedPath = NormalizeLocalImagePath(imagePath);
            if (string.IsNullOrWhiteSpace(normalizedPath))
            {
                return false;
            }

            if (normalizedPath.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
                || normalizedPath.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
                || normalizedPath.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var fileName = Path.GetFileNameWithoutExtension(normalizedPath)?.Trim();
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return false;
            }

            if (Regex.IsMatch(fileName, @"\.00[12]$", RegexOptions.IgnoreCase))
            {
                return true;
            }

            var docTitle = Path.GetFileNameWithoutExtension(currentDocPath)?.Trim();
            if (string.IsNullOrWhiteSpace(docTitle))
            {
                return false;
            }

            return fileName.StartsWith($"{docTitle}.00", StringComparison.OrdinalIgnoreCase);
        }

        private static void TryDeleteImageByPath(string imagePath, string currentDocPath)
        {
            try
            {
                var normalizedPath = NormalizeLocalImagePath(imagePath);
                if (string.IsNullOrWhiteSpace(normalizedPath))
                {
                    return;
                }

                if (normalizedPath.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
                    || normalizedPath.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
                    || normalizedPath.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                var absolutePath = Path.IsPathRooted(normalizedPath)
                    ? normalizedPath
                    : Path.GetFullPath(Path.Combine(Path.GetDirectoryName(currentDocPath) ?? ".", normalizedPath.Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar)));

                if (File.Exists(absolutePath))
                {
                    File.Delete(absolutePath);
                }
            }
            catch
            {
            }
        }

        private static void RemoveLeadingTitleLine(List<string> lines, bool[] removeFlags, string currentDocPath)
        {
            var docTitle = Path.GetFileNameWithoutExtension(currentDocPath)?.Trim();
            if (string.IsNullOrWhiteSpace(docTitle))
            {
                return;
            }

            for (var i = 0; i < lines.Count && i < 30; i++)
            {
                if (removeFlags[i])
                {
                    continue;
                }

                var line = lines[i];
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                var normalizedLine = NormalizeMarkdownText(line);
                if (string.Equals(normalizedLine, docTitle, StringComparison.OrdinalIgnoreCase))
                {
                    removeFlags[i] = true;
                }

                return;
            }
        }

        private static string NormalizeMarkdownText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            var result = text.Trim();
            result = Regex.Replace(result, @"^#{1,6}\s*", string.Empty);
            result = result.Trim('*', '_', '`', ' ');
            return result.Trim();
        }

        private static string NormalizeLocalImagePath(string imagePath)
        {
            if (string.IsNullOrWhiteSpace(imagePath))
            {
                return null;
            }

            var normalized = imagePath.Trim().Trim('<', '>');
            try
            {
                normalized = Uri.UnescapeDataString(normalized);
            }
            catch
            {
            }

            return normalized;
        }
    }
}
