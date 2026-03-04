using System.Text.Json.Serialization;

namespace feishu_doc_export.Dtos
{
    /// <summary>
    /// Docx 文档块
    /// </summary>
    public class DocxBlockItemDto
    {
        [JsonPropertyName("block_id")]
        public string BlockId { get; set; }

        [JsonPropertyName("block_type")]
        public int BlockType { get; set; }

        [JsonPropertyName("parent_id")]
        public string ParentId { get; set; }

        public DocxFileBlockDto File { get; set; }
    }

    /// <summary>
    /// 文件块信息（block_type=23）
    /// </summary>
    public class DocxFileBlockDto
    {
        public string Name { get; set; }

        public string Token { get; set; }
    }
}
