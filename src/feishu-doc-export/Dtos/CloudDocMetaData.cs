using System.Text.Json.Serialization;

namespace feishu_doc_export.Dtos
{
    /// <summary>
    /// 云文档详情接口返回结构（兼容 data.file 与 data 直出两种形态）
    /// </summary>
    public class CloudDocMetaData
    {
        [JsonPropertyName("file")]
        public CloudDocDto File { get; set; }

        public string Name { get; set; }

        [JsonPropertyName("parent_token")]
        public string ParentToken { get; set; }

        public string Token { get; set; }

        public string Type { get; set; }

        public string Url { get; set; }

        [JsonPropertyName("has_child")]
        public bool HasChild { get; set; }
    }
}
