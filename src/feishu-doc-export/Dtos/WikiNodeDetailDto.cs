using System.Text.Json.Serialization;

namespace feishu_doc_export.Dtos
{
    public class WikiNodeDetailDto
    {
        [JsonPropertyName("node")]
        public WikiNodeItemDto Node { get; set; }
    }
}
