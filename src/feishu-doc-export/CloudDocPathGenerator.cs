using feishu_doc_export.Dtos;
using feishu_doc_export.Helper;

namespace feishu_doc_export
{
    public static class CloudDocPathGenerator
    {
        /// <summary>
        /// 文档token和路径的映射
        /// </summary>
        private static Dictionary<string, string> documentPaths;

        private static HashSet<string> generatedPaths;

        public static void GenerateDocumentPaths(List<CloudDocDto> documents, string rootFolderPath)
        {
            documentPaths = new Dictionary<string, string>();
            generatedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var topDocuments = documents.Where(x =>
                string.IsNullOrWhiteSpace(x.ParentToken) || !documents.Any(d => d.Token == x.ParentToken));
            foreach (var document in topDocuments)
            {
                if (!documentPaths.ContainsKey(document.Token))
                {
                    GenerateDocumentPath(document, rootFolderPath, documents);
                }
            }
        }

        private static void GenerateDocumentPath(CloudDocDto document, string parentFolderPath, List<CloudDocDto> documents)
        {
            var name = PathNameHelper.SanitizePathSegment(document.Name);
            var documentFolderPath = Path.Combine(parentFolderPath, name);
            documentFolderPath = EnsureUniquePath(documentFolderPath, document.Token);

            documentPaths[document.Token] = documentFolderPath;
            generatedPaths.Add(documentFolderPath);

            foreach (var childDocument in GetChildDocuments(document, documents))
            {
                GenerateDocumentPath(childDocument, documentFolderPath, documents);
            }
        }

        private static string EnsureUniquePath(string path, string token)
        {
            if (!generatedPaths.Contains(path))
            {
                return path;
            }

            var suffix = string.IsNullOrWhiteSpace(token)
                ? Guid.NewGuid().ToString("N")[..6]
                : token.Length <= 6 ? token : token[^6..];

            var index = 0;
            string candidate;
            do
            {
                candidate = index == 0 ? $"{path}_{suffix}" : $"{path}_{suffix}_{index}";
                index++;
            } while (generatedPaths.Contains(candidate));

            return candidate;
        }

        private static IEnumerable<CloudDocDto> GetChildDocuments(CloudDocDto document, List<CloudDocDto> documents)
        {
            return documents.Where(d => d.ParentToken == document.Token);
        }

        /// <summary>
        /// 获取文档的存储路径
        /// </summary>
        public static string GetDocumentPath(string objToken)
        {
            documentPaths.TryGetValue(objToken, out string path);
            return path;
        }
    }
}
