using feishu_doc_export.Dtos;
using feishu_doc_export.Helper;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using WebApiClientCore.Exceptions;

namespace feishu_doc_export.HttpApi
{
    public interface IFeiShuHttpApiCaller
    {
        #region 知识空间Wiki
        /// <summary>
        /// 获取空间子节点列表
        /// </summary>
        /// <param name="spaceId">知识空间Id</param>
        /// <param name="pageToken">分页token，第一次查询没有</param>
        /// <param name="parentNodeToken">父节点token</param>
        /// <returns></returns>
        Task<PagedResult<WikiNodeItemDto>> GetWikiNodeList(string spaceId, string pageToken = null, string parentNodeToken = null);

        /// <summary>
        /// 获取文档块列表
        /// </summary>
        /// <param name="documentId"></param>
        /// <param name="pageToken"></param>
        /// <returns></returns>
        Task<PagedResult<DocxBlockItemDto>> GetDocxBlocks(string documentId, string pageToken = null);

        /// <summary>
        /// 获取知识空间下全部文档节点
        /// </summary>
        /// <param name="spaceId"></param>
        /// <returns></returns>
        Task<List<WikiNodeItemDto>> GetAllWikiNode(string spaceId);

        /// <summary>
        /// 根据token获取节点信息
        /// </summary>
        /// <param name="token"></param>
        /// <returns></returns>
        Task<WikiNodeItemDto> GetWikiNodeInfo(string token);

        /// <summary>
        /// 根据根节点token递归获取整棵子树
        /// </summary>
        /// <param name="rootToken"></param>
        /// <returns></returns>
        Task<List<WikiNodeItemDto>> GetWikiSubTreeByToken(string rootToken);

        /// <summary>
        /// 递归获取知识空间下指定节点下的所有子节点（包括孙节点）
        /// </summary>
        /// <param name="spaceId">知识空间id</param>
        /// <param name="parentNodeToken">父节点token</param>
        /// <returns></returns>
        Task<List<WikiNodeItemDto>> GetWikiChildNode(string spaceId, string parentNodeToken);

        /// <summary>
        /// 获取所有的知识库
        /// </summary>
        /// <returns></returns>
        Task<PagedResult<WikiSpaceDto>> GetWikiSpaces();

        /// <summary>
        /// 获取知识库详细信息
        /// </summary>
        /// <param name="spaceId"></param>
        /// <returns></returns>
        Task<WikiSpaceInfo> GetWikiSpaceInfo(string spaceId);
        #endregion

        #region 下载文档
        /// <summary>
        /// 创建导出任务
        /// </summary>
        /// <param name="fileExtension">导出文件扩展名</param>
        /// <param name="token">文档token</param>
        /// <param name="type">导出文档类型</param>
        /// <returns></returns>
        Task<ExportOutputDto> CreateExportTask(string fileExtension, string token, string type);

        /// <summary>
        /// 查询导出任务的结果
        /// </summary>
        /// <param name="ticket"></param>
        /// <param name="token"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        Task<ExportTaskResultDto> QueryExportTaskResult(string ticket, string token);

        /// <summary>
        /// 下载文档文件
        /// </summary>
        /// <param name="fileToken"></param>
        /// <returns></returns>
        Task<byte[]> DownLoad(string fileToken);

        /// <summary>
        /// 下载文件
        /// </summary>
        /// <param name="fileToken"></param>
        /// <returns></returns>
        Task<byte[]> DownLoadFile(string fileToken);

        /// <summary>
        /// 下载文件（静默，不打印错误日志）
        /// </summary>
        /// <param name="fileToken"></param>
        /// <returns></returns>
        Task<byte[]> DownLoadFileSilently(string fileToken);

        /// <summary>
        /// 下载媒体文件（静默，不打印错误日志）
        /// </summary>
        /// <param name="fileToken"></param>
        /// <returns></returns>
        Task<byte[]> DownLoadMediaSilently(string fileToken);

        /// <summary>
        /// 下载附件（优先files接口，403/404时自动回退media接口）
        /// </summary>
        /// <param name="fileToken"></param>
        /// <returns></returns>
        Task<byte[]> DownLoadAttachmentBestEffort(string fileToken);
        #endregion

        #region 个人空间云文档
        /// <summary>
        /// 获取文件夹信息
        /// </summary>
        /// <param name="folderToken"></param>
        /// <returns></returns>
        Task<CloudDocFolderMeta> GetFolderMeta(string folderToken);

        /// <summary>
        /// 获取云文档节点详情
        /// </summary>
        /// <param name="fileToken"></param>
        /// <returns></returns>
        Task<CloudDocDto> GetCloudDocMeta(string fileToken);
        /// <summary>
        /// 获取个人空间云文档
        /// </summary>
        /// <param name="folderToken"></param>
        /// <param name="pageToken"></param>
        /// <returns></returns>
        Task<PagedResult<CloudDocDto>> GetCloudDocList(string folderToken = null, string pageToken = null);

        /// <summary>
        /// 获取个人空间云文档指定文件夹下的所有文档
        /// </summary>
        /// <param name="folderToken"></param>
        /// <returns></returns>
        Task<List<CloudDocDto>> GetFolderAllCloudDoc(string folderToken);

        /// <summary>
        /// 获取指定节点及全部后代
        /// </summary>
        /// <param name="rootToken"></param>
        /// <param name="includeRoot"></param>
        /// <returns></returns>
        Task<List<CloudDocDto>> GetCloudDocTree(string rootToken, bool includeRoot = true);

        /// <summary>
        /// 递归获取子文档
        /// </summary>
        /// <param name="parentNodeToken"></param>
        /// <returns></returns>
        Task<List<CloudDocDto>> GetChildCloudDoc(string parentNodeToken);
        #endregion
    }
    public class FeiShuHttpApiCaller : IFeiShuHttpApiCaller
    {
        private readonly IFeiShuHttpApi _feiShuHttpApi;
        private static string OpenApiHost => FeiShuConsts.GetOpenApiEndPoint(GlobalConfig.Platform);

        public FeiShuHttpApiCaller(IFeiShuHttpApi feiShuHttpApi)
        {
            _feiShuHttpApi = feiShuHttpApi;
        }

        private static string BuildOpenApiUrl(string relativePath)
        {
            return FeiShuConsts.BuildOpenApiUrl(relativePath, GlobalConfig.Platform);
        }

        #region 获取知识库所有的文档节点

        public async Task<PagedResult<WikiNodeItemDto>> GetWikiNodeList(string spaceId, string pageToken = null, string parentNodeToken = null)
        {
            StringBuilder urlBuilder = new StringBuilder($"{OpenApiHost}/open-apis/wiki/v2/spaces/{spaceId}/nodes?page_size=50");// page_size=50
            if (!string.IsNullOrWhiteSpace(pageToken))
            {
                urlBuilder.Append($"&page_token={pageToken}");
            }

            if (!string.IsNullOrWhiteSpace(parentNodeToken))
            {
                urlBuilder.Append($"&parent_node_token={parentNodeToken}");
            }

            var resultData = await _feiShuHttpApi.GetWikeNodeList(urlBuilder.ToString());

            return resultData.Data;
        }

        public async Task<PagedResult<DocxBlockItemDto>> GetDocxBlocks(string documentId, string pageToken = null)
        {
            StringBuilder urlBuilder = new StringBuilder($"{OpenApiHost}/open-apis/docx/v1/documents/{documentId}/blocks?page_size=50");
            if (!string.IsNullOrWhiteSpace(pageToken))
            {
                urlBuilder.Append($"&page_token={pageToken}");
            }

            var resultData = await _feiShuHttpApi.GetDocxBlocks(urlBuilder.ToString());
            return resultData.Data;
        }

        public async Task<List<WikiNodeItemDto>> GetAllWikiNode(string spaceId)
        {
            try
            {
                List<WikiNodeItemDto> nodes = new List<WikiNodeItemDto>();
                string pageToken = null;
                bool hasMore;
                do
                {
                    // 分页获取顶级节点，pageToken = null时为获取第一页
                    var pagedResult = await GetWikiNodeList(spaceId, pageToken);
                    nodes.AddRange(pagedResult.Items);

                    foreach (var item in pagedResult.Items)
                    {
                        if (item.HasChild)
                        {
                            List<WikiNodeItemDto> childNodes = await GetWikiChildNode(spaceId, item.NodeToken);
                            nodes.AddRange(childNodes);
                        }
                    }

                    pageToken = pagedResult.PageToken;
                    hasMore = pagedResult.HasMore;

                } while (hasMore && !string.IsNullOrWhiteSpace(pageToken));

                return nodes;

            }
            catch (HttpRequestException ex)
            {
                var responseData = string.Empty;

                if (ex.InnerException is ApiResponseStatusException statusException)
                {
                    var response = statusException.ResponseMessage;

                    // 响应的数据
                    responseData = await response.Content.ReadAsStringAsync();
                }

                LogHelper.LogErrorExit($"请求异常！！！\r\n 异常信息： {ex.Message}，\r\n 堆栈信息： {ex.StackTrace}，\r\n 响应信息：{responseData} \r\n");
                throw;
            }
        }

        public async Task<WikiNodeItemDto> GetWikiNodeInfo(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return null;
            }

            try
            {
                var url = BuildOpenApiUrl($"/open-apis/wiki/v2/spaces/get_node?token={token}");
                var res = await _feiShuHttpApi.GetWikiNodeInfo(url);
                return res?.Data?.Node;
            }
            catch (HttpRequestException ex)
            {
                var responseData = string.Empty;

                if (ex.InnerException is ApiResponseStatusException statusException)
                {
                    var response = statusException.ResponseMessage;
                    responseData = await response.Content.ReadAsStringAsync();
                }

                LogHelper.LogError($"根据Token获取知识库节点失败，Token:{token}，异常信息：{ex.Message}，响应信息：{responseData}");
                throw;
            }
        }

        public async Task<List<WikiNodeItemDto>> GetWikiSubTreeByToken(string rootToken)
        {
            if (string.IsNullOrWhiteSpace(rootToken))
            {
                return new List<WikiNodeItemDto>();
            }

            var rootNode = await GetWikiNodeInfo(rootToken);
            if (rootNode == null)
            {
                return new List<WikiNodeItemDto>();
            }

            var result = new List<WikiNodeItemDto> { rootNode };
            var visited = new HashSet<string>(StringComparer.Ordinal);
            var queue = new Queue<string>();
            visited.Add(rootNode.NodeToken);
            queue.Enqueue(rootNode.NodeToken);

            while (queue.Any())
            {
                var parentNodeToken = queue.Dequeue();
                string pageToken = null;
                bool hasMore;

                do
                {
                    var pagedResult = await GetWikiNodeList(rootNode.SpaceId, pageToken, parentNodeToken);
                    var children = pagedResult?.Items ?? new List<WikiNodeItemDto>();
                    foreach (var child in children)
                    {
                        if (string.IsNullOrWhiteSpace(child?.NodeToken) || !visited.Add(child.NodeToken))
                        {
                            continue;
                        }

                        result.Add(child);
                        if (child.HasChild)
                        {
                            queue.Enqueue(child.NodeToken);
                        }
                    }

                    pageToken = pagedResult?.PageToken;
                    hasMore = pagedResult?.HasMore ?? false;
                } while (hasMore && !string.IsNullOrWhiteSpace(pageToken));
            }

            return result;
        }

        public async Task<List<WikiNodeItemDto>> GetWikiChildNode(string spaceId, string parentNodeToken)
        {
            List<WikiNodeItemDto> childNodes = new List<WikiNodeItemDto>();
            string pageToken = null;
            bool hasMore;
            do
            {
                var pagedResult = await GetWikiNodeList(spaceId, pageToken, parentNodeToken);
                childNodes.AddRange(pagedResult.Items);

                foreach (var item in pagedResult.Items)
                {
                    if (item.HasChild)
                    {
                        List<WikiNodeItemDto> grandChildNodes = await GetWikiChildNode(spaceId, item.NodeToken);
                        childNodes.AddRange(grandChildNodes);
                    }
                }

                pageToken = pagedResult.PageToken;
                hasMore = pagedResult.HasMore;

            } while (hasMore && !string.IsNullOrWhiteSpace(pageToken));

            return childNodes;
        }

        #endregion

        #region 导出文档
        public async Task<byte[]> DownLoadFile(string fileToken)
        {
            try
            {
                var result = await _feiShuHttpApi.DownLoadFile(BuildOpenApiUrl($"/open-apis/drive/v1/files/{fileToken}/download"));

                return result;
            }
            catch (HttpRequestException ex)
            {
                LogHelper.LogError($"请求异常！！！\r\n 异常信息： {ex.Message}，\r\n 堆栈信息： {ex.StackTrace} \r\n");
                throw;
            }
        }

        public async Task<ExportOutputDto> CreateExportTask(string fileExtension, string token, string type)
        {
            var request = RequestData.CreateExportTask(fileExtension, token, type);

            try
            {
                var result = await _feiShuHttpApi.CreateExportTask(BuildOpenApiUrl("/open-apis/drive/v1/export_tasks"), request);
                return result.Data;
            }
            catch (HttpRequestException ex) when (ex.InnerException is ApiResponseStatusException statusException)
            {
                // 响应状态码异常
                var response = statusException.ResponseMessage;

                // 响应的数据
                var responseData = await response.Content.ReadAsStringAsync();

                if (responseData.Contains("1069902"))
                {
                    string message = $"无阅读或导出权限，已忽略，请手动下载。飞书服务端响应数据为：{responseData}";
                    throw new CustomException(message, 1069902);
                }
            }

            return null;
        }

        public async Task<ExportTaskResultDto> QueryExportTaskResult(string ticket, string token)
        {
            int status;// 0成功，1初始化，2处理中
            const int maxPollCount = 600; // 600 * 300ms ~= 180s

            var data = new ExportTaskResultDto();
            var pollCount = 0;
            do
            {
                pollCount++;
                if (pollCount > maxPollCount)
                {
                    throw new TimeoutException($"导出任务轮询超时，ticket={ticket} token={token}");
                }

                var result = await _feiShuHttpApi.QueryExportTask(BuildOpenApiUrl($"/open-apis/drive/v1/export_tasks/{ticket}?token={token}"));

                status = result.Data.Result.JobStatus;

                switch (status)
                {
                    case 0:
                        data = result.Data;
                        break;
                    case 1:
                    case 2:
                        await Task.Delay(300);
                        break;
                    default:
                        throw new Exception($"Error: {result.Data.Result.JobErrorMsg}，ErrorCode:{status}");
                }

            } while (status != 0);

            return data;
        }

        public async Task<byte[]> DownLoad(string fileToken)
        {
            var result = await _feiShuHttpApi.DownLoad(BuildOpenApiUrl($"/open-apis/drive/v1/export_tasks/file/{fileToken}/download"));

            return result;
        }

        #endregion

        #region 知识库
        public async Task<PagedResult<WikiSpaceDto>> GetWikiSpaces()
        {
            try
            {
                var res = await _feiShuHttpApi.GetWikiSpaces(BuildOpenApiUrl("/open-apis/wiki/v2/spaces"));

                return res.Data;
            }
            catch (HttpRequestException ex)
            {
                var responseData = string.Empty;

                if (ex.InnerException is ApiResponseStatusException statusException)
                {
                    var response = statusException.ResponseMessage;

                    // 响应的数据
                    responseData = await response.Content.ReadAsStringAsync();
                }

                LogHelper.LogErrorExit($"请求异常！！！\r\n 异常信息： {ex.Message}，\r\n 堆栈信息： {ex.StackTrace}，\r\n 响应信息：{responseData} \r\n");
                throw;
            }
        }

        public async Task<WikiSpaceInfo> GetWikiSpaceInfo(string spaceId)
        {
            try
            {
                var res = await _feiShuHttpApi.GetWikiSpaceInfo(BuildOpenApiUrl($"/open-apis/wiki/v2/spaces/{spaceId}"));

                return res.Data;
            }
            catch (HttpRequestException ex)
            {
                var responseData = string.Empty;

                if (ex.InnerException is ApiResponseStatusException statusException)
                {
                    var response = statusException.ResponseMessage;

                    // 响应的数据
                    responseData = await response.Content.ReadAsStringAsync();
                }

                LogHelper.LogErrorExit($"请求异常！！！\r\n 异常信息： {ex.Message}，\r\n 堆栈信息： {ex.StackTrace}，\r\n 响应信息：{responseData} \r\n");
                throw;
            }
        }
        #endregion

        #region 个人空间云文档
        public async Task<CloudDocFolderMeta> GetFolderMeta(string folderToken)
        {
            try
            {
                var res = await _feiShuHttpApi.GetFolderMeta(BuildOpenApiUrl($"/open-apis/drive/explorer/v2/folder/{folderToken}/meta"));

                return res.Data;
            }
            catch (HttpRequestException ex)
            {
                var responseData = string.Empty;

                if (ex.InnerException is ApiResponseStatusException statusException)
                {
                    var response = statusException.ResponseMessage;

                    // 响应的数据
                    responseData = await response.Content.ReadAsStringAsync();
                }

                LogHelper.LogErrorExit($"请求异常！！！\r\n 异常信息： {ex.Message}，\r\n 堆栈信息： {ex.StackTrace}，\r\n 响应信息：{responseData} \r\n");
                throw;
            }
        }

        public async Task<byte[]> DownLoadFileSilently(string fileToken)
        {
            var result = await _feiShuHttpApi.DownLoadFile(BuildOpenApiUrl($"/open-apis/drive/v1/files/{fileToken}/download"));
            return result;
        }

        public async Task<byte[]> DownLoadMediaSilently(string fileToken)
        {
            var result = await _feiShuHttpApi.DownLoadMedia(BuildOpenApiUrl($"/open-apis/drive/v1/medias/{fileToken}/download"));
            return result;
        }

        public async Task<byte[]> DownLoadAttachmentBestEffort(string fileToken)
        {
            try
            {
                return await DownLoadFileSilently(fileToken);
            }
            catch (HttpRequestException ex) when (ex.InnerException is ApiResponseStatusException statusException
                && (statusException.ResponseMessage?.StatusCode == HttpStatusCode.Forbidden
                    || statusException.ResponseMessage?.StatusCode == HttpStatusCode.NotFound))
            {
                return await DownLoadMediaSilently(fileToken);
            }
        }

        public async Task<CloudDocDto> GetCloudDocMeta(string fileToken)
        {
            try
            {
                var res = await _feiShuHttpApi.GetCloudDocMeta(BuildOpenApiUrl($"/open-apis/drive/v1/files/{fileToken}"));
                if (res?.Data == null)
                {
                    return null;
                }

                if (res.Data.File != null)
                {
                    if (string.IsNullOrWhiteSpace(res.Data.File.Token))
                    {
                        res.Data.File.Token = fileToken;
                    }

                    return res.Data.File;
                }

                return new CloudDocDto
                {
                    Token = string.IsNullOrWhiteSpace(res.Data.Token) ? fileToken : res.Data.Token,
                    Name = res.Data.Name,
                    ParentToken = res.Data.ParentToken,
                    Type = res.Data.Type,
                    Url = res.Data.Url,
                    HasChild = res.Data.HasChild
                };
            }
            catch
            {
                return null;
            }
        }

        public async Task<PagedResult<CloudDocDto>> GetCloudDocList(string folderToken = null, string pageToken = null)
        {
            StringBuilder urlBuilder = new StringBuilder($"{OpenApiHost}/open-apis/drive/v1/files?folder_token={folderToken}&page_size=50");// page_size=50
            if (!string.IsNullOrWhiteSpace(pageToken))
            {
                urlBuilder.Append($"&page_token={pageToken}");
            }

            var resultData = await _feiShuHttpApi.GetCloudDocList(urlBuilder.ToString());

            return resultData.Data;
        }

        public async Task<List<CloudDocDto>> GetFolderAllCloudDoc(string folderToken)
        {
            return await GetCloudDocTree(folderToken, false);
        }

        public async Task<List<CloudDocDto>> GetCloudDocTree(string rootToken, bool includeRoot = true)
        {
            try
            {
                var nodes = new List<CloudDocDto>();
                var visitedItemTokens = new HashSet<string>();
                var visitedParentTokens = new HashSet<string>();

                if (includeRoot)
                {
                    var rootNode = await GetCloudDocMeta(rootToken);
                    if (rootNode != null)
                    {
                        if (string.IsNullOrWhiteSpace(rootNode.Token))
                        {
                            rootNode.Token = rootToken;
                        }

                        if (string.IsNullOrWhiteSpace(rootNode.Name))
                        {
                            rootNode.Name = rootToken;
                        }
                    }

                    if (rootNode == null)
                    {
                        try
                        {
                            var folderMeta = await GetFolderMeta(rootToken);
                            rootNode = new CloudDocDto
                            {
                                Token = rootToken,
                                Name = folderMeta?.Name ?? rootToken,
                                Type = "folder"
                            };
                        }
                        catch
                        {
                            // ignore
                        }
                    }

                    if (rootNode != null && visitedItemTokens.Add(rootNode.Token))
                    {
                        nodes.Add(rootNode);
                    }

                    if (rootNode != null && !CanContainChildren(rootNode))
                    {
                        return nodes;
                    }
                }

                List<CloudDocDto> childNodes = await GetChildCloudDoc(rootToken, visitedItemTokens, visitedParentTokens);
                nodes.AddRange(childNodes);

                return nodes;

            }
            catch (HttpRequestException ex)
            {
                var responseData = string.Empty;

                if (ex.InnerException is ApiResponseStatusException statusException)
                {
                    var response = statusException.ResponseMessage;

                    // 响应的数据
                    responseData = await response.Content.ReadAsStringAsync();
                }

                LogHelper.LogErrorExit($"请求异常！！！\r\n 异常信息： {ex.Message}，\r\n 堆栈信息： {ex.StackTrace}，\r\n 响应信息：{responseData} \r\n");
                throw;
            }
        }

        public async Task<List<CloudDocDto>> GetChildCloudDoc(string parentNodeToken)
        {
            return await GetChildCloudDoc(parentNodeToken, new HashSet<string>(), new HashSet<string>());
        }

        private async Task<List<CloudDocDto>> GetChildCloudDoc(string parentNodeToken, HashSet<string> visitedItemTokens, HashSet<string> visitedParentTokens)
        {
            List<CloudDocDto> childNodes = new List<CloudDocDto>();

            if (!visitedParentTokens.Add(parentNodeToken))
            {
                return childNodes;
            }

            string pageToken = null;
            bool hasMore;
            do
            {
                PagedResult<CloudDocDto> pagedResult;
                try
                {
                    pagedResult = await GetCloudDocList(parentNodeToken, pageToken);
                }
                catch (Exception ex)
                {
                    LogHelper.LogWarn($"读取云文档子节点失败，父Token:{parentNodeToken}，错误:{ex.Message}");
                    break;
                }

                if (pagedResult?.Files == null || !pagedResult.Files.Any())
                {
                    break;
                }

                foreach (var item in pagedResult.Files)
                {
                    if (visitedItemTokens.Add(item.Token))
                    {
                        childNodes.Add(item);
                    }
                }

                foreach (var item in pagedResult.Files)
                {
                    if (CanContainChildren(item))
                    {
                        List<CloudDocDto> grandChildNodes = await GetChildCloudDoc(item.Token, visitedItemTokens, visitedParentTokens);
                        childNodes.AddRange(grandChildNodes);
                    }
                }

                pageToken = pagedResult.PageToken ?? pagedResult.NextPageToken;
                hasMore = pagedResult.HasMore;

            } while (hasMore && !string.IsNullOrWhiteSpace(pageToken));

            return childNodes;
        }

        private static bool CanContainChildren(CloudDocDto item)
        {
            if (item == null)
            {
                return false;
            }

            if (string.Equals(item.Type, "file", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (item.HasChild)
            {
                return true;
            }

            return string.Equals(item.Type, "folder", StringComparison.OrdinalIgnoreCase);
        }

        #endregion

    }
}
