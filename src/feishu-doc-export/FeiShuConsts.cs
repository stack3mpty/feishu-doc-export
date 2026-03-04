using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace feishu_doc_export
{
    public static class FeiShuConsts
    {
        /// <summary>
        /// 国内飞书平台标识
        /// </summary>
        public const string PlatformFeishu = "feishu";

        /// <summary>
        /// 海外Lark平台标识
        /// </summary>
        public const string PlatformLark = "lark";

        /// <summary>
        /// 飞书（国内）OpenAPI地址
        /// </summary>
        public const string FeishuOpenApiEndPoint = "https://open.feishu.cn";

        /// <summary>
        /// Lark（海外）OpenAPI地址
        /// </summary>
        public const string LarkOpenApiEndPoint = "https://open.larksuite.com";

        public static string GetOpenApiEndPoint(string platform)
        {
            if (string.Equals(platform, PlatformLark, StringComparison.OrdinalIgnoreCase))
            {
                return LarkOpenApiEndPoint;
            }

            return FeishuOpenApiEndPoint;
        }

        public static string BuildOpenApiUrl(string relativePath, string platform)
        {
            if (string.IsNullOrWhiteSpace(relativePath))
            {
                return GetOpenApiEndPoint(platform);
            }

            if (Uri.TryCreate(relativePath, UriKind.Absolute, out _))
            {
                return relativePath;
            }

            var path = relativePath.StartsWith("/", StringComparison.Ordinal) ? relativePath : "/" + relativePath;
            return $"{GetOpenApiEndPoint(platform)}{path}";
        }
    }
}
