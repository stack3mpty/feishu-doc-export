using feishu_doc_export.Dtos;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebApiClientCore;
using WebApiClientCore.Attributes;

namespace feishu_doc_export.HttpApi
{
    [HttpHost(FeiShuConsts.FeishuOpenApiEndPoint)]
    public interface IFeiShuHttpApi : IHttpApi
    {
        /// <summary>
        /// 获取自建应用的AccessToken
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost]
        Task<AccessTokenDto> GetTenantAccessToken([Uri] string url, [JsonContent] object request);

        /// <summary>
        /// 获取所有的知识库
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [OAuthToken]
        [JsonReturn]
        Task<ResponseData<PagedResult<WikiSpaceDto>>> GetWikiSpaces([Uri] string url);

        /// <summary>
        /// 获取知识库详细信息
        /// </summary>
        /// <param name="spaceId"></param>
        /// <returns></returns>
        [HttpGet]
        [OAuthToken]
        [JsonReturn]
        Task<ResponseData<WikiSpaceInfo>> GetWikiSpaceInfo([Uri] string url);

        /// <summary>
        /// 根据token获取知识库节点信息
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        [HttpGet]
        [OAuthToken]
        [JsonReturn]
        Task<ResponseData<WikiNodeDetailDto>> GetWikiNodeInfo([Uri] string url);

        /// <summary>
        /// 获取知识空间子节点列表
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [OAuthToken]
        [JsonReturn]
        Task<ResponseData<PagedResult<WikiNodeItemDto>>> GetWikeNodeList([Uri] string url);

        /// <summary>
        /// 获取文档块列表
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        [HttpGet]
        [OAuthToken]
        [JsonReturn]
        Task<ResponseData<PagedResult<DocxBlockItemDto>>> GetDocxBlocks([Uri] string url);

        /// <summary>
        /// 获取个人空间指定文件夹下的文档列表
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        [HttpGet]
        [OAuthToken]
        [JsonReturn]
        Task<ResponseData<PagedResult<CloudDocDto>>> GetCloudDocList([Uri] string url);

        [HttpGet]
        [OAuthToken]
        [JsonReturn]
        Task<ResponseData<CloudDocFolderMeta>> GetFolderMeta([Uri] string url);

        /// <summary>
        /// 获取云文档文件详情
        /// </summary>
        /// <param name="fileToken"></param>
        /// <returns></returns>
        [HttpGet]
        [OAuthToken]
        [JsonReturn]
        Task<ResponseData<CloudDocMetaData>> GetCloudDocMeta([Uri] string url);

        /// <summary>
        /// 创建文档导出任务结果
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost]
        [OAuthToken]
        [JsonReturn]
        Task<ResponseData<ExportOutputDto>> CreateExportTask([Uri] string url, [JsonContent] object request);

        /// <summary>
        /// 查询导出任务
        /// </summary>
        /// <param name="ticket"></param>
        /// <param name="token"></param>
        /// <returns></returns>
        [HttpGet]
        [OAuthToken]
        [JsonReturn]
        Task<ResponseData<ExportTaskResultDto>> QueryExportTask([Uri] string url);

        /// <summary>
        /// 下载导出文件
        /// </summary>
        /// <param name="token"></param>
        /// <returns></returns>
        [HttpGet]
        [OAuthToken]
        Task<byte[]> DownLoad([Uri] string url);

        /// <summary>
        /// 下载文件
        /// </summary>
        /// <param name="fileToken"></param>
        /// <returns></returns>
        [HttpGet]
        [OAuthToken]
        Task<byte[]> DownLoadFile([Uri] string url);

        /// <summary>
        /// 下载媒体文件（部分文档附件需通过该接口下载）
        /// </summary>
        /// <param name="fileToken"></param>
        /// <returns></returns>
        [HttpGet]
        [OAuthToken]
        Task<byte[]> DownLoadMedia([Uri] string url);
    }
}
