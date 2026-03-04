namespace feishu_doc_export.Dtos
{
    public class ExportDocumentManifestRoot
    {
        public string ManifestVersion { get; set; } = "1.0";

        public DateTimeOffset GeneratedAtUtc { get; set; }

        public string SourceType { get; set; }

        public string SaveType { get; set; }

        public string WikiSpaceId { get; set; }

        public string RootToken { get; set; }

        public string ExportPath { get; set; }

        public List<ExportDocumentManifestItem> Documents { get; set; } = new List<ExportDocumentManifestItem>();
    }

    public class ExportDocumentManifestItem
    {
        public string DocumentToken { get; set; }

        public string NodeToken { get; set; }

        public string ParentDocumentToken { get; set; }

        public string ParentNodeToken { get; set; }

        public string DocumentType { get; set; }

        public string Title { get; set; }

        public string RelativeOutputPath { get; set; }

        public string Status { get; set; }

        public string Error { get; set; }

        public DateTimeOffset UpdatedAtUtc { get; set; }
    }

    public class ExportAttachmentManifestItem
    {
        public string DocumentToken { get; set; }

        public string DocumentRelativePath { get; set; }

        public string BlockId { get; set; }

        public int BlockIndex { get; set; }

        public int BlockType { get; set; }

        public string FileToken { get; set; }

        public string FileName { get; set; }

        public string RelativeOutputPath { get; set; }

        public string Sha256 { get; set; }

        public string Status { get; set; }

        public string Error { get; set; }

        public DateTimeOffset UpdatedAtUtc { get; set; }
    }

    public class AttachmentFailureRecord
    {
        public string FailureKey { get; set; }

        public string DocumentToken { get; set; }

        public string DocumentRelativePath { get; set; }

        public string BlockId { get; set; }

        public int BlockIndex { get; set; }

        public int BlockType { get; set; }

        public string FileToken { get; set; }

        public string FileName { get; set; }

        public string Error { get; set; }

        public int FailureCount { get; set; }

        public DateTimeOffset LastFailureAtUtc { get; set; }
    }
}
