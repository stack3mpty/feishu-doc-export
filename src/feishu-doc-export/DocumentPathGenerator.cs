using feishu_doc_export.Dtos;
using feishu_doc_export.Helper;

namespace feishu_doc_export
{
    public static class DocumentPathGenerator
    {
        /// <summary>
        /// 文档objToken和路径的映射
        /// </summary>
        private static Dictionary<string, string> documentPaths;

        /// <summary>
        /// 文档nodeToken和路径的映射
        /// </summary>
        private static Dictionary<string, string> documentPaths2;

        private static HashSet<string> generatedPaths;

        public static void GenerateDocumentPaths(List<WikiNodeItemDto> documents, string rootFolderPath)
        {
            documentPaths = new Dictionary<string, string>();
            documentPaths2 = new Dictionary<string, string>();
            generatedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var topDocument = documents.Where(x =>
                string.IsNullOrWhiteSpace(x.ParentNodeToken) || !documents.Any(d => d.NodeToken == x.ParentNodeToken));
            foreach (var document in topDocument)
            {
                GenerateDocumentPath(document, rootFolderPath, documents);
            }
        }

        private static void GenerateDocumentPath(WikiNodeItemDto document, string parentFolderPath, List<WikiNodeItemDto> documents)
        {
            var title = PathNameHelper.SanitizePathSegment(document.Title);
            var documentFolderPath = Path.Combine(parentFolderPath, title);
            documentFolderPath = EnsureUniquePath(documentFolderPath, document.ObjToken);

            documentPaths[document.ObjToken] = documentFolderPath;
            documentPaths2[document.NodeToken] = documentFolderPath;
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

        private static IEnumerable<WikiNodeItemDto> GetChildDocuments(WikiNodeItemDto document, List<WikiNodeItemDto> documents)
        {
            return documents.Where(d => d.ParentNodeToken == document.NodeToken);
        }

        /// <summary>
        /// 获取文档的存储路径
        /// </summary>
        public static string GetDocumentPath(string objToken)
        {
            if (documentPaths == null || string.IsNullOrWhiteSpace(objToken))
            {
                return null;
            }

            documentPaths.TryGetValue(objToken, out string path);
            return path;
        }

        /// <summary>
        /// 获取文档的存储路径
        /// </summary>
        public static string GetDocumentPathByNodeToken(string nodeToken)
        {
            if (documentPaths2 == null || string.IsNullOrWhiteSpace(nodeToken))
            {
                return null;
            }

            documentPaths2.TryGetValue(nodeToken, out string path);
            return path;
        }
    }
}
