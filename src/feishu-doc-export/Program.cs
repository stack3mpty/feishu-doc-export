
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using feishu_doc_export.Dtos;
using feishu_doc_export.Helper;
using feishu_doc_export.HttpApi;
using Microsoft.Extensions.DependencyInjection;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using System.Threading;
using WebApiClientCore;
using WebApiClientCore.Exceptions;

namespace feishu_doc_export
{
    internal class Program
    {
        static IFeiShuHttpApiCaller feiShuApiCaller;
        static ExportProgressStore exportProgressStore;
        static Dictionary<string, AttachmentFailureRecord> attachmentFailureByKey = new Dictionary<string, AttachmentFailureRecord>(StringComparer.Ordinal);
        static Dictionary<string, ExportDocumentManifestItem> documentManifestByToken = new Dictionary<string, ExportDocumentManifestItem>(StringComparer.Ordinal);
        static Dictionary<string, ExportAttachmentManifestItem> attachmentManifestByKey = new Dictionary<string, ExportAttachmentManifestItem>(StringComparer.Ordinal);
        static Stopwatch runStopwatch = new Stopwatch();
        static int totalDocumentCount;
        static int completedDocumentCount;
        static int discoveredAttachmentCount;
        static int completedAttachmentCount;
        static int successAttachmentCount;
        static int skippedAttachmentCount;
        static int dedupAttachmentCount;
        static int failedAttachmentCount;
        static CancellationTokenSource heartbeatCts;
        static Task heartbeatTask;
        const int HeartbeatIntervalSeconds = 10;
        static string heartbeatCurrentDocument = "-";
        static string heartbeatCurrentAttachment = "-";

        static async Task Main(string[] args)
        {
            GlobalConfig.Init(args);
            LogHelper.LogInfo($"当前平台：{GlobalConfig.Platform}，OpenAPI域名：{FeiShuConsts.GetOpenApiEndPoint(GlobalConfig.Platform)}");
            LogHelper.LogInfo($"当前鉴权模式：{GlobalConfig.AuthMode}");

            if (!Directory.Exists(GlobalConfig.ExportPath))
            {
                LogHelper.LogWarnExit($"指定的导出目录({GlobalConfig.ExportPath})不存在！！！");
            }

            IOC.Init();
            feiShuApiCaller = IOC.IoContainer.GetService<IFeiShuHttpApiCaller>();

            Stopwatch stopwatch = new Stopwatch();

            // 不支持导出的文件
            List<string> noSupportExportFiles = new List<string>();
            attachmentFailureByKey = new Dictionary<string, AttachmentFailureRecord>(StringComparer.Ordinal);
            documentManifestByToken = new Dictionary<string, ExportDocumentManifestItem>(StringComparer.Ordinal);
            attachmentManifestByKey = new Dictionary<string, ExportAttachmentManifestItem>(StringComparer.Ordinal);
            totalDocumentCount = 0;
            completedDocumentCount = 0;
            discoveredAttachmentCount = 0;
            completedAttachmentCount = 0;
            successAttachmentCount = 0;
            skippedAttachmentCount = 0;
            dedupAttachmentCount = 0;
            failedAttachmentCount = 0;
            runStopwatch.Reset();
            heartbeatCts = null;
            heartbeatTask = null;
            heartbeatCurrentDocument = "-";
            heartbeatCurrentAttachment = "-";
            exportProgressStore = new ExportProgressStore(GlobalConfig.ExportPath);
            LogHelper.LogInfo($"断点续传状态文件：{exportProgressStore.StatePath}（启用状态：{GlobalConfig.Resume}）");
            LogHelper.LogInfo($"快速续跑策略：已完成文档{(GlobalConfig.ResumeSkipAttachmentScan ? "不再扫描附件块" : "继续扫描附件块")}");
            if (GlobalConfig.Resume)
            {
                LoadAttachmentFailureRecords();
                LogHelper.LogInfo($"已加载附件失败记录：{attachmentFailureByKey.Count}");
            }

            if (string.Equals(GlobalConfig.Type, "cloudDoc", StringComparison.OrdinalIgnoreCase))
            {

                if (string.IsNullOrWhiteSpace(GlobalConfig.CloudDocFolder))
                {
                    LogHelper.LogWarnExit("导出对象为个人空间云文档时，请填写【folderToken】参数（支持文件夹或父文档token）");
                }

                var cloudRoot = await feiShuApiCaller.GetCloudDocMeta(GlobalConfig.CloudDocFolder);
                var rootName = cloudRoot?.Name;
                if (string.IsNullOrWhiteSpace(rootName))
                {
                    try
                    {
                        var folderMeta = await feiShuApiCaller.GetFolderMeta(GlobalConfig.CloudDocFolder);
                        rootName = folderMeta?.Name;
                    }
                    catch
                    {
                        rootName = GlobalConfig.CloudDocFolder;
                    }
                }

                if (string.IsNullOrWhiteSpace(rootName))
                {
                    rootName = GlobalConfig.CloudDocFolder;
                }

                Console.WriteLine($"正在加载个人空间云文档节点【{rootName}】及其后代所有文档信息，请耐心等待...");

                // 获取指定节点（父节点）及其所有后代文档
                var selfDocs = await feiShuApiCaller.GetCloudDocTree(GlobalConfig.CloudDocFolder);
                if (!selfDocs.Any())
                {
                    LogHelper.LogWarnExit($"未找到节点【{GlobalConfig.CloudDocFolder}】下可导出的文档");
                }

                var exportTargets = selfDocs
                    .Where(x => !string.Equals(x.Type, "folder", StringComparison.OrdinalIgnoreCase))
                    .ToList();
                InitRunProgress(exportTargets.Count);
                stopwatch.Start();
                runStopwatch.Start();
                StartHeartbeat();

                // 文档路径映射字典
                CloudDocPathGenerator.GenerateDocumentPaths(selfDocs, GlobalConfig.ExportPath);

                // 记录导出的文档数量
                int count = 1;
                foreach (var item in exportTargets)
                {
                    heartbeatCurrentDocument = item.Name ?? item.Token;

                    UpsertDocumentManifest(new ExportDocumentManifestItem
                    {
                        DocumentToken = item.Token,
                        ParentDocumentToken = item.ParentToken,
                        DocumentType = item.Type,
                        Title = item.Name,
                        Status = "pending"
                    });

                    var isSupport = GlobalConfig.GetFileExtension(item.Type, out string fileExt);

                    // 如果该文件类型不支持导出
                    if (!isSupport)
                    {
                        noSupportExportFiles.Add(item.Name);
                        UpdateDocumentManifestStatus(item.Token, "unsupported", null, null);
                        LogHelper.LogWarn($"文档【{item.Name}】不支持导出，已忽略。如有需要请手动下载。");
                        ReportDocumentProgress(item.Name, "unsupported");
                        continue;
                    }

                    // 文档为文件类型则直接下载文件
                    if (fileExt == "file")
                    {
                        try
                        {
                            var filePath = GetFileOutputPath(item.Token, item.Name);
                            if (await TrySkipExistingDocumentAsync(item.Token, item.Type, filePath))
                            {
                                UpdateDocumentManifestStatus(item.Token, "skipped_existing", filePath, null);
                                ReportDocumentProgress(item.Name, "skipped_existing");
                                continue;
                            }

                            Console.WriteLine($"正在导出文档————————{count++}.【{item.Name}】");
                            await DownLoadFile(item.Token, item.Name, filePath);
                            UpdateDocumentManifestStatus(item.Token, "success", filePath, null);
                            ReportDocumentProgress(item.Name, "success");

                            continue;
                        }
                        catch (HttpRequestException ex)
                        {
                            noSupportExportFiles.Add(item.Name);
                            UpdateDocumentManifestStatus(item.Token, "failed", filePath: null, error: ex.Message);
                            LogHelper.LogError($"下载文档【{item.Name}】时出现请求异常！！！异常信息：{ex.Message}，堆栈信息：{ex.StackTrace}");
                            ReportDocumentProgress(item.Name, "failed");
                            continue;
                        }
                        catch (Exception ex)
                        {
                            noSupportExportFiles.Add(item.Name);
                            UpdateDocumentManifestStatus(item.Token, "failed", filePath: null, error: ex.Message);
                            LogHelper.LogError($"下载文档【{item.Name}】时出现未知异常，已忽略。请手动下载。异常信息：{ex.Message}，堆栈信息：{ex.StackTrace}");
                            ReportDocumentProgress(item.Name, "failed");
                            continue;
                        }
                    }

                    // 用于展示的文件后缀名称
                    var showFileExt = fileExt;
                    // 用于指定文件下载类型
                    var fileExtension = fileExt;

                    // 只有当飞书文档类型为docx时才支持使用自定义文档保存类型
                    if (fileExt == "docx")
                    {
                        showFileExt = GlobalConfig.DocSaveType;

                        if (GlobalConfig.DocSaveType == "pdf")
                        {
                            fileExtension = GlobalConfig.DocSaveType;
                        }
                    }

                    // 文件名超出长度限制，不支持导出
                    if (item.Name.Length > 64)
                    {
                        var left64FileName = item.Name.PadLeft(61) + $"···.{fileExt}";
                        noSupportExportFiles.Add($"(文件名超长){left64FileName}");
                        UpdateDocumentManifestStatus(item.Token, "unsupported", null, "文件名超长");
                        Console.WriteLine($"文档【{left64FileName}】的文件命名长度超出系统文件命名的长度限制，已忽略");
                        ReportDocumentProgress(item.Name, "unsupported");
                        continue;
                    }

                    var documentOutputPath = GetDocumentOutputPath(item.Token, fileExtension);
                    if (await TrySkipExistingDocumentAsync(item.Token, item.Type, documentOutputPath))
                    {
                        UpdateDocumentManifestStatus(item.Token, "skipped_existing", documentOutputPath, null);
                        ReportDocumentProgress(item.Name, "skipped_existing");
                        continue;
                    }

                    Console.WriteLine($"正在导出文档————————{count++}.【{item.Name}.{showFileExt}】");

                    try
                    {
                        await DownLoadDocument(fileExtension, item.Token, item.Type, documentOutputPath);
                        UpdateDocumentManifestStatus(item.Token, "success", documentOutputPath, null);
                        ReportDocumentProgress(item.Name, "success");
                    }
                    catch (HttpRequestException ex)
                    {
                        noSupportExportFiles.Add(item.Name);
                        UpdateDocumentManifestStatus(item.Token, "failed", documentOutputPath, ex.Message);
                        LogHelper.LogError($"下载文档【{item.Name}】时出现请求异常，异常信息：{ex.Message}，堆栈信息：{ex.StackTrace}");
                        ReportDocumentProgress(item.Name, "failed");
                    }
                    catch (CustomException ex)
                    {
                        noSupportExportFiles.Add(item.Name);
                        UpdateDocumentManifestStatus(item.Token, "failed", documentOutputPath, ex.Message);
                        LogHelper.LogWarn($"文档【{item.Name}】{ex.Message}");
                        ReportDocumentProgress(item.Name, "failed");
                    }
                    catch (Exception ex)
                    {
                        noSupportExportFiles.Add(item.Name);
                        UpdateDocumentManifestStatus(item.Token, "failed", documentOutputPath, ex.Message);
                        LogHelper.LogError($"下载文档【{item.Name}】时出现未知异常，已忽略，请手动下载。异常信息：{ex.Message}，堆栈信息：{ex.StackTrace}");
                        ReportDocumentProgress(item.Name, "failed");
                    }
                }
            }
            else
            {
                List<WikiNodeItemDto> wikiNodes;
                if (string.Equals(GlobalConfig.Type, "wikiToken", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrWhiteSpace(GlobalConfig.WikiRootNodeToken))
                    {
                        LogHelper.LogWarnExit("wikiToken模式下，必须填写--rootNodeToken（或--rootToken）。");
                    }

                    var rootNode = await feiShuApiCaller.GetWikiNodeInfo(GlobalConfig.WikiRootNodeToken);
                    if (rootNode == null)
                    {
                        LogHelper.LogWarnExit($"无法根据Token【{GlobalConfig.WikiRootNodeToken}】获取知识库节点信息，请确认token有效且应用具备访问权限。");
                    }

                    GlobalConfig.WikiRootNodeToken = rootNode.NodeToken;
                    GlobalConfig.WikiSpaceId = rootNode.SpaceId;
                    Console.WriteLine($"正在加载知识库【{GlobalConfig.WikiSpaceId}】中父文档【{rootNode.Title}】及其后代所有文档信息，请耐心等待...");
                    wikiNodes = await feiShuApiCaller.GetWikiSubTreeByToken(rootNode.NodeToken);
                    if (!wikiNodes.Any())
                    {
                        LogHelper.LogWarnExit($"未找到父文档节点【{GlobalConfig.WikiRootNodeToken}】或其下无可导出文档");
                    }
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(GlobalConfig.WikiSpaceId))
                    {
                        var wikiSpaces = await feiShuApiCaller.GetWikiSpaces();
                        var wikiSpaceDict = wikiSpaces.Items
                            .Select((x, i) => new { Index = i + 1, WikiSpace = x })
                            .ToDictionary(x => x.Index, x => x.WikiSpace);

                        if (wikiSpaceDict.Any())
                        {
                            Console.WriteLine($"以下是所有支持导出的知识库：");

                            foreach (var item in wikiSpaceDict)
                            {
                                Console.WriteLine($"【{item.Key}.】{item.Value.Name}");
                            }
                            Console.WriteLine("请选择知识库（输入知识库的序号）：");
                            var index = int.Parse(Console.ReadLine());
                            GlobalConfig.WikiSpaceId = wikiSpaceDict[index].Spaceid;
                        }
                        else
                        {
                            LogHelper.LogWarnExit("没有可支持导出的知识库！！！");
                        }
                    }

                    var wikiSpaceInfo = await feiShuApiCaller.GetWikiSpaceInfo(GlobalConfig.WikiSpaceId);
                    Console.WriteLine($"正在加载知识库【{wikiSpaceInfo.Space.Name}】的所有文档信息，请耐心等待...");

                    // 获取知识库下的所有文档
                    wikiNodes = await feiShuApiCaller.GetAllWikiNode(GlobalConfig.WikiSpaceId);
                    if (!string.IsNullOrWhiteSpace(GlobalConfig.WikiRootNodeToken))
                    {
                        wikiNodes = GetWikiSubTree(wikiNodes, GlobalConfig.WikiRootNodeToken);

                        if (!wikiNodes.Any())
                        {
                            LogHelper.LogWarnExit($"未找到父文档节点【{GlobalConfig.WikiRootNodeToken}】或其下无可导出文档");
                        }
                    }
                }

                InitRunProgress(wikiNodes.Count);
                stopwatch.Start();
                runStopwatch.Start();
                StartHeartbeat();

                // 文档路径映射字典
                DocumentPathGenerator.GenerateDocumentPaths(wikiNodes, GlobalConfig.ExportPath);
                var nodeTokenToObjToken = wikiNodes.ToDictionary(x => x.NodeToken, x => x.ObjToken, StringComparer.Ordinal);

                // 记录导出的文档数量
                int count = 1;
                foreach (var item in wikiNodes)
                {
                    heartbeatCurrentDocument = item.Title ?? item.ObjToken;
                    nodeTokenToObjToken.TryGetValue(item.ParentNodeToken ?? string.Empty, out var parentObjToken);
                    UpsertDocumentManifest(new ExportDocumentManifestItem
                    {
                        DocumentToken = item.ObjToken,
                        NodeToken = item.NodeToken,
                        ParentNodeToken = item.ParentNodeToken,
                        ParentDocumentToken = parentObjToken,
                        DocumentType = item.ObjType,
                        Title = item.Title,
                        Status = "pending"
                    });

                    var isSupport = GlobalConfig.GetFileExtension(item.ObjType, out string fileExt);

                    // 如果该文件类型不支持导出
                    if (!isSupport)
                    {
                        noSupportExportFiles.Add(item.Title);
                        UpdateDocumentManifestStatus(item.ObjToken, "unsupported", null, null);
                        LogHelper.LogWarn($"文档【{item.Title}】不支持导出，已忽略。如有需要请手动下载。");
                        ReportDocumentProgress(item.Title, "unsupported");
                        continue;
                    }

                    // 文档为文件类型则直接下载文件
                    if (fileExt == "file")
                    {
                        try
                        {
                            var filePath = GetFileOutputPath(item.ObjToken, item.Title);
                            if (await TrySkipExistingDocumentAsync(item.ObjToken, item.ObjType, filePath))
                            {
                                UpdateDocumentManifestStatus(item.ObjToken, "skipped_existing", filePath, null);
                                ReportDocumentProgress(item.Title, "skipped_existing");
                                continue;
                            }

                            Console.WriteLine($"正在导出文档————————{count++}.【{item.Title}】");
                            await DownLoadFile(item.ObjToken, item.Title, filePath);
                            UpdateDocumentManifestStatus(item.ObjToken, "success", filePath, null);
                            ReportDocumentProgress(item.Title, "success");

                            continue;
                        }
                        catch (HttpRequestException ex)
                        {
                            noSupportExportFiles.Add(item.Title);
                            UpdateDocumentManifestStatus(item.ObjToken, "failed", null, ex.Message);
                            LogHelper.LogError($"下载文档【{item.Title}】时出现请求异常！！！异常信息：{ex.Message}，堆栈信息：{ex.StackTrace}");
                            ReportDocumentProgress(item.Title, "failed");
                            continue;
                        }
                        catch (Exception ex)
                        {
                            noSupportExportFiles.Add(item.Title);
                            UpdateDocumentManifestStatus(item.ObjToken, "failed", null, ex.Message);
                            LogHelper.LogWarn($"下载文档【{item.Title}】时出现未知异常，已忽略。请手动下载。异常信息：{ex.Message}");
                            ReportDocumentProgress(item.Title, "failed");
                            continue;
                        }
                    }

                    // 用于展示的文件后缀名称
                    var showFileExt = fileExt;
                    // 用于指定文件下载类型
                    var fileExtension = fileExt;

                    // 只有当飞书文档类型为docx时才支持使用自定义文档保存类型
                    if (fileExt == "docx")
                    {
                        showFileExt = GlobalConfig.DocSaveType;

                        if (GlobalConfig.DocSaveType == "pdf")
                        {
                            fileExtension = GlobalConfig.DocSaveType;
                        }
                    }

                    // 文件名超出长度限制，不支持导出
                    if (item.Title.Length > 64)
                    {
                        var left64FileName = item.Title.PadLeft(61) + $"···.{fileExt}";
                        noSupportExportFiles.Add($"(文件名超长){left64FileName}");
                        UpdateDocumentManifestStatus(item.ObjToken, "unsupported", null, "文件名超长");
                        Console.WriteLine($"文档【{left64FileName}】的文件命名长度超出系统文件命名的长度限制，已忽略");
                        ReportDocumentProgress(item.Title, "unsupported");
                        continue;
                    }

                    var documentOutputPath = GetDocumentOutputPath(item.ObjToken, fileExtension);
                    if (await TrySkipExistingDocumentAsync(item.ObjToken, item.ObjType, documentOutputPath))
                    {
                        UpdateDocumentManifestStatus(item.ObjToken, "skipped_existing", documentOutputPath, null);
                        ReportDocumentProgress(item.Title, "skipped_existing");
                        continue;
                    }

                    Console.WriteLine($"正在导出文档————————{count++}.【{item.Title}.{showFileExt}】");

                    try
                    {
                        await DownLoadDocument(fileExtension, item.ObjToken, item.ObjType, documentOutputPath);
                        UpdateDocumentManifestStatus(item.ObjToken, "success", documentOutputPath, null);
                        ReportDocumentProgress(item.Title, "success");
                    }
                    catch (HttpRequestException ex)
                    {
                        noSupportExportFiles.Add(item.Title);
                        UpdateDocumentManifestStatus(item.ObjToken, "failed", documentOutputPath, ex.Message);
                        LogHelper.LogError($"下载文档【{item.Title}】时出现请求异常！！！异常信息：{ex.Message}，堆栈信息：{ex.StackTrace}");
                        ReportDocumentProgress(item.Title, "failed");
                    }
                    catch (CustomException ex)
                    {
                        noSupportExportFiles.Add(item.Title);
                        UpdateDocumentManifestStatus(item.ObjToken, "failed", documentOutputPath, ex.Message);
                        LogHelper.LogWarn($"文档【{item.Title}】{ex.Message}");
                        ReportDocumentProgress(item.Title, "failed");
                    }
                    catch (Exception ex)
                    {
                        noSupportExportFiles.Add(item.Title);
                        UpdateDocumentManifestStatus(item.ObjToken, "failed", documentOutputPath, ex.Message);
                        LogHelper.LogError($"下载文档【{item.Title}】时出现未知异常，已忽略，请手动下载。异常信息：{ex.Message}，堆栈信息：{ex.StackTrace}");
                        ReportDocumentProgress(item.Title, "failed");
                    }
                }
            }

            await RetryPendingAttachmentFailuresAsync();
            await StopHeartbeatAsync();
            Console.WriteLine("—————————————————————————————文档已全部导出—————————————————————————————");
            Console.WriteLine(GetProgressSnapshotLine("completed"));
            Console.WriteLine(noSupportExportFiles.Any() ? "以下是所有无法导出的文档（包含不支持导出、导出异常的文档）" : "");

            // 输出不支持导入的文档
            for (int i = 0; i < noSupportExportFiles.Count; i++)
            {
                Console.WriteLine($"{i + 1}.【{noSupportExportFiles[i]}】");
            }

            await WriteAttachmentFailureFilesAsync();

            await WriteManifestFilesAsync();

            stopwatch.Stop();
            TimeSpan elapsedTime = stopwatch.Elapsed;
            // 输出执行时间（以秒为单位）
            double seconds = elapsedTime.TotalSeconds;

            if (GlobalConfig.Quit)
            {
                Console.WriteLine($"程序执行结束，总耗时{seconds}（秒）。已自动退出程序！");
                return;
            }

            Console.WriteLine($"程序执行结束，总耗时{seconds}（秒）。请按任意键退出！");
            if (!Console.IsInputRedirected)
            {
                try
                {
                    Console.ReadKey();
                }
                catch
                {
                }
            }
        }

        /// <summary>
        /// 下载文档到本地
        /// </summary>
        /// <param name="fileExtension">文档导出的文件类型（docx）</param>
        /// <param name="objToken"></param>
        /// <param name="type"></param>
        /// <param name="outputPath"></param>
        /// <returns></returns>
        static async Task DownLoadDocument(string fileExtension, string objToken, string type, string outputPath)
        {
            var exportTaskDto = await ExecuteWithRetryAsync(() => feiShuApiCaller.CreateExportTask(fileExtension, objToken, type), "创建导出任务", objToken);

            if (exportTaskDto == null)
            {
                return;
            }

            var exportTaskResult = await ExecuteWithRetryAsync(() => feiShuApiCaller.QueryExportTaskResult(exportTaskDto.Ticket, objToken), "查询导出任务", objToken);
            var taskInfo = exportTaskResult?.Result;

            if (taskInfo == null)
            {
                throw new Exception($"文档Token:{objToken}，导出任务未返回有效结果。");
            }

            if (!string.Equals(taskInfo.JobErrorMsg, "success", StringComparison.OrdinalIgnoreCase))
            {
                throw new Exception($"文档Token:{objToken}，导出任务失败：{taskInfo.JobErrorMsg}");
            }

            var bytes = await ExecuteWithRetryAsync(() => feiShuApiCaller.DownLoad(taskInfo.FileToken), "下载导出文件", taskInfo.FileToken);

            if (string.Equals(type, "docx", StringComparison.OrdinalIgnoreCase))
            {
                await DownLoadDocAttachments(objToken, outputPath);
            }

            if (fileExtension == "docx" && GlobalConfig.DocSaveType == "md")
            {
                await SaveToMarkdownFile(bytes, outputPath);
                exportProgressStore?.MarkDocumentCompleted(objToken, outputPath);
                return;
            }

            await outputPath.Save(bytes);
            exportProgressStore?.MarkDocumentCompleted(objToken, outputPath);
        }

        /// <summary>
        /// 下载文件到本地
        /// </summary>
        /// <param name="objToken"></param>
        /// <param name="sourceFileName"></param>
        /// <returns></returns>
        static async Task DownLoadFile(string objToken, string sourceFileName = null, string outputPath = null)
        {
            var bytes = await ExecuteWithRetryAsync(() => feiShuApiCaller.DownLoadAttachmentBestEffort(objToken), "下载文件", objToken);
            var filePath = outputPath ?? GetFileOutputPath(objToken, sourceFileName);
            await filePath.Save(bytes);
            exportProgressStore?.MarkDocumentCompleted(objToken, filePath);
        }

        /// <summary>
        /// 保存为Markdown文件
        /// </summary>
        /// <param name="bytes"></param>
        /// <param name="fileSavePath"></param>
        static async Task SaveToMarkdownFile(byte[] bytes,string fileSavePath)
        {
            using (MemoryStream stream = new MemoryStream(bytes))
            {
                // 加载 Word 文档
                Document doc = new Document(stream);

                // 遍历文档中的所有形状（包括图片）
                foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
                {
                    if (shape.HasImage)
                    {
                        // 清空图片描述
                        shape.AlternativeText = "";
                    }
                }

                // 创建Markdown保存选项
                MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
                // 文件保存的文件夹路径
                var saveDirPath = Path.GetDirectoryName(fileSavePath);
                // 设置文章中图片的存储路径
                saveOptions.ImagesFolder = Path.Combine(saveDirPath, "images");
                // 重构文件名
                var fileName = Path.GetFileNameWithoutExtension(fileSavePath) + ".md";
                // 文件最终的保存路径
                var mdFileSavePath = Path.Combine(saveDirPath, fileName);
                doc.Save(mdFileSavePath, saveOptions);

                // 处理 Markdown 文件，替换图片和文档的引用路径为相对路径
                var markdownContent = await File.ReadAllTextAsync(mdFileSavePath);
                var replacedContent = markdownContent
                    .ReplaceImagePath(mdFileSavePath)
                    .ReplaceDocRefPath(mdFileSavePath)
                    .CleanupExportArtifacts(mdFileSavePath);
                await File.WriteAllTextAsync(mdFileSavePath, replacedContent);
            }

        }

        static List<WikiNodeItemDto> GetWikiSubTree(List<WikiNodeItemDto> allNodes, string rootNodeToken)
        {
            var root = allNodes.FirstOrDefault(x => x.NodeToken == rootNodeToken);
            if (root == null)
            {
                return new List<WikiNodeItemDto>();
            }

            var result = new List<WikiNodeItemDto>();
            var queue = new Queue<string>();
            var visited = new HashSet<string>();

            queue.Enqueue(rootNodeToken);
            while (queue.Any())
            {
                var currentNodeToken = queue.Dequeue();
                if (!visited.Add(currentNodeToken))
                {
                    continue;
                }

                var current = allNodes.FirstOrDefault(x => x.NodeToken == currentNodeToken);
                if (current != null)
                {
                    result.Add(current);
                }

                var children = allNodes.Where(x => x.ParentNodeToken == currentNodeToken);
                foreach (var child in children)
                {
                    queue.Enqueue(child.NodeToken);
                }
            }

            return result;
        }

        static async Task DownLoadDocAttachments(string documentToken, string fileSavePath)
        {
            var downloadDir = Path.Combine(GlobalConfig.ExportPath, "attachments", "by-token");
            var documentRelativePath = ToRelativePath(fileSavePath);
            var tokenToOutputPath = new Dictionary<string, string>(StringComparer.Ordinal);
            string pageToken = null;
            bool hasMore;
            int blockIndex = 0;
            do
            {
                PagedResult<DocxBlockItemDto> pagedResult;
                try
                {
                    pagedResult = await ExecuteWithRetryAsync(() => feiShuApiCaller.GetDocxBlocks(documentToken, pageToken), "读取文档块", documentToken);
                }
                catch (Exception ex)
                {
                    RecordAttachmentFailure(documentToken, "文档块读取失败", "N/A", ex.Message, blockId: "N/A", blockIndex: 0, blockType: 23, documentRelativePath: documentRelativePath);
                    return;
                }

                var blocks = pagedResult?.Items ?? new List<DocxBlockItemDto>();
                foreach (var block in blocks)
                {
                    blockIndex++;
                    if (block?.BlockType != 23 || block.File == null || string.IsNullOrWhiteSpace(block.File.Token))
                    {
                        continue;
                    }

                    var fileToken = block.File.Token;
                    var suggestedFileName = string.IsNullOrWhiteSpace(block.File.Name) ? fileToken : block.File.Name;
                    heartbeatCurrentAttachment = suggestedFileName;
                    discoveredAttachmentCount++;
                    if (tokenToOutputPath.TryGetValue(fileToken, out var knownOutputPath))
                    {
                        UpsertAttachmentManifest(new ExportAttachmentManifestItem
                        {
                            DocumentToken = documentToken,
                            DocumentRelativePath = documentRelativePath,
                            BlockId = block.BlockId,
                            BlockIndex = blockIndex,
                            BlockType = block.BlockType,
                            FileToken = fileToken,
                            FileName = suggestedFileName,
                            RelativeOutputPath = ToRelativePath(knownOutputPath),
                            Sha256 = ComputeSha256FromFile(knownOutputPath),
                            Status = "dedup_in_document",
                            UpdatedAtUtc = DateTimeOffset.UtcNow
                        });
                        RemoveAttachmentFailure(documentToken, block.BlockId, fileToken, suggestedFileName);
                        ReportAttachmentProgress(documentToken, suggestedFileName, "dedup_in_document");
                        continue;
                    }

                    if (GlobalConfig.Resume && exportProgressStore.TryGetDownloadedAttachmentPath(fileToken, out var existingPath))
                    {
                        UpsertAttachmentManifest(new ExportAttachmentManifestItem
                        {
                            DocumentToken = documentToken,
                            DocumentRelativePath = documentRelativePath,
                            BlockId = block.BlockId,
                            BlockIndex = blockIndex,
                            BlockType = block.BlockType,
                            FileToken = fileToken,
                            FileName = suggestedFileName,
                            RelativeOutputPath = ToRelativePath(existingPath),
                            Sha256 = ComputeSha256FromFile(existingPath),
                            Status = "skipped_existing",
                            UpdatedAtUtc = DateTimeOffset.UtcNow
                        });
                        tokenToOutputPath[fileToken] = existingPath;
                        RemoveAttachmentFailure(documentToken, block.BlockId, fileToken, suggestedFileName);
                        ReportAttachmentProgress(documentToken, suggestedFileName, "skipped_existing");
                        continue;
                    }

                    try
                    {
                        var attachmentBytes = await ExecuteWithRetryAsync(() => feiShuApiCaller.DownLoadAttachmentBestEffort(fileToken), "下载附件", fileToken);
                        var savePath = BuildAttachmentSavePath(downloadDir, fileToken, suggestedFileName);
                        await savePath.Save(attachmentBytes);
                        exportProgressStore?.MarkAttachmentCompleted(fileToken, savePath);
                        tokenToOutputPath[fileToken] = savePath;
                        UpsertAttachmentManifest(new ExportAttachmentManifestItem
                        {
                            DocumentToken = documentToken,
                            DocumentRelativePath = documentRelativePath,
                            BlockId = block.BlockId,
                            BlockIndex = blockIndex,
                            BlockType = block.BlockType,
                            FileToken = fileToken,
                            FileName = suggestedFileName,
                            RelativeOutputPath = ToRelativePath(savePath),
                            Sha256 = ComputeSha256(attachmentBytes),
                            Status = "success",
                            UpdatedAtUtc = DateTimeOffset.UtcNow
                        });
                        RemoveAttachmentFailure(documentToken, block.BlockId, fileToken, suggestedFileName);
                        ReportAttachmentProgress(documentToken, suggestedFileName, "success");
                    }
                    catch (Exception ex)
                    {
                        UpsertAttachmentManifest(new ExportAttachmentManifestItem
                        {
                            DocumentToken = documentToken,
                            DocumentRelativePath = documentRelativePath,
                            BlockId = block.BlockId,
                            BlockIndex = blockIndex,
                            BlockType = block.BlockType,
                            FileToken = fileToken,
                            FileName = suggestedFileName,
                            Status = "failed",
                            Error = ex.Message,
                            UpdatedAtUtc = DateTimeOffset.UtcNow
                        });
                        ReportAttachmentProgress(documentToken, suggestedFileName, "failed");
                        RecordAttachmentFailure(documentToken, suggestedFileName, fileToken, ex.Message, block.BlockId, blockIndex, block.BlockType, documentRelativePath);
                    }
                }

                pageToken = pagedResult?.PageToken ?? pagedResult?.NextPageToken;
                hasMore = pagedResult?.HasMore ?? false;
            } while (hasMore && !string.IsNullOrWhiteSpace(pageToken));

            // 当前文档附件块读取完成后，清理历史的文档级读取失败记录。
            RemoveAttachmentFailure(documentToken, "N/A", "N/A", "文档块读取失败");
        }

        static async Task<bool> TrySkipExistingDocumentAsync(string documentToken, string docType, string outputPath)
        {
            if (!GlobalConfig.Resume || exportProgressStore == null)
            {
                return false;
            }

            if (!exportProgressStore.ShouldSkipDocument(documentToken, outputPath))
            {
                return false;
            }

            Console.WriteLine($"跳过已完成文档（断点续传）：{Path.GetFileName(outputPath)}");
            if (string.Equals(docType, "docx", StringComparison.OrdinalIgnoreCase) && !GlobalConfig.ResumeSkipAttachmentScan)
            {
                // 兼容旧行为：仅在显式关闭快速续跑时，已完成文档仍尝试补齐附件。
                await DownLoadDocAttachments(documentToken, outputPath);
            }

            return true;
        }

        static string GetDocumentBasePath(string objToken)
        {
            return string.Equals(GlobalConfig.Type, "cloudDoc", StringComparison.OrdinalIgnoreCase)
                ? CloudDocPathGenerator.GetDocumentPath(objToken)
                : DocumentPathGenerator.GetDocumentPath(objToken);
        }

        static string GetDocumentOutputPath(string objToken, string fileExtension)
        {
            var filePath = GetDocumentBasePath(objToken) + "." + fileExtension;
            if (string.Equals(fileExtension, "docx", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(GlobalConfig.DocSaveType, "md", StringComparison.OrdinalIgnoreCase))
            {
                return Path.ChangeExtension(filePath, ".md");
            }

            return filePath;
        }

        static string GetFileOutputPath(string objToken, string sourceFileName)
        {
            var filePath = GetDocumentBasePath(objToken);
            return EnsureFilePathWithExtension(filePath, sourceFileName);
        }

        static void InitRunProgress(int totalDocs)
        {
            totalDocumentCount = Math.Max(totalDocs, 0);
            completedDocumentCount = 0;
            discoveredAttachmentCount = 0;
            completedAttachmentCount = 0;
            successAttachmentCount = 0;
            skippedAttachmentCount = 0;
            dedupAttachmentCount = 0;
            failedAttachmentCount = 0;
            heartbeatCurrentDocument = "-";
            heartbeatCurrentAttachment = "-";

            Console.WriteLine($"[导出进度] 初始化完成：待处理文档 {totalDocumentCount} 篇。");
        }

        static void StartHeartbeat()
        {
            if (heartbeatTask != null)
            {
                return;
            }

            heartbeatCts = new CancellationTokenSource();
            heartbeatTask = Task.Run(async () =>
            {
                while (!heartbeatCts.Token.IsCancellationRequested)
                {
                    try
                    {
                        await Task.Delay(TimeSpan.FromSeconds(HeartbeatIntervalSeconds), heartbeatCts.Token);
                    }
                    catch (TaskCanceledException)
                    {
                        break;
                    }

                    if (heartbeatCts.Token.IsCancellationRequested)
                    {
                        break;
                    }

                    Console.WriteLine(GetProgressSnapshotLine("heartbeat"));
                }
            }, heartbeatCts.Token);
        }

        static async Task StopHeartbeatAsync()
        {
            if (heartbeatTask == null || heartbeatCts == null)
            {
                return;
            }

            try
            {
                heartbeatCts.Cancel();
                await heartbeatTask;
            }
            catch (TaskCanceledException)
            {
            }
            finally
            {
                heartbeatCts.Dispose();
                heartbeatCts = null;
                heartbeatTask = null;
            }
        }

        static void ReportDocumentProgress(string title, string status)
        {
            completedDocumentCount++;
            if (completedDocumentCount > totalDocumentCount)
            {
                completedDocumentCount = totalDocumentCount;
            }
            heartbeatCurrentDocument = string.IsNullOrWhiteSpace(title) ? heartbeatCurrentDocument : title;
            heartbeatCurrentAttachment = "-";

            var percent = totalDocumentCount <= 0
                ? 100d
                : completedDocumentCount * 100d / totalDocumentCount;

            var elapsed = runStopwatch.Elapsed;
            var eta = completedDocumentCount <= 0
                ? TimeSpan.Zero
                : TimeSpan.FromSeconds(Math.Max(0, elapsed.TotalSeconds / completedDocumentCount * (totalDocumentCount - completedDocumentCount)));

            Console.WriteLine(
                $"[导出进度] 文档 {completedDocumentCount}/{totalDocumentCount} ({percent:F1}%) 状态={status} 当前={title} 已用={FormatDuration(elapsed)} 预计剩余={FormatDuration(eta)}");
        }

        static void ReportAttachmentProgress(string documentToken, string fileName, string status)
        {
            completedAttachmentCount++;
            switch (status)
            {
                case "success":
                    successAttachmentCount++;
                    break;
                case "skipped_existing":
                    skippedAttachmentCount++;
                    break;
                case "dedup_in_document":
                    dedupAttachmentCount++;
                    break;
                case "failed":
                    failedAttachmentCount++;
                    break;
            }

            Console.WriteLine(
                $"[附件进度] 已处理={completedAttachmentCount} 已发现={discoveredAttachmentCount} 成功={successAttachmentCount} 跳过={skippedAttachmentCount} 去重={dedupAttachmentCount} 失败={failedAttachmentCount} 当前文档Token={documentToken} 文件={fileName} 状态={status}");
            if (!string.IsNullOrWhiteSpace(fileName))
            {
                heartbeatCurrentAttachment = fileName;
            }
        }

        static string GetProgressSnapshotLine(string stage)
        {
            var percent = totalDocumentCount <= 0
                ? 100d
                : completedDocumentCount * 100d / totalDocumentCount;
            return
                $"[导出汇总] 阶段={stage} 文档={completedDocumentCount}/{totalDocumentCount} ({percent:F1}%) 当前文档={heartbeatCurrentDocument} 当前附件={heartbeatCurrentAttachment} 附件已处理={completedAttachmentCount}（成功={successAttachmentCount} 跳过={skippedAttachmentCount} 去重={dedupAttachmentCount} 失败={failedAttachmentCount}） 总耗时={FormatDuration(runStopwatch.Elapsed)}";
        }

        static string FormatDuration(TimeSpan duration)
        {
            if (duration.TotalSeconds < 60)
            {
                return $"{duration.TotalSeconds:F1}s";
            }

            return $"{(int)duration.TotalMinutes}m{duration.Seconds:D2}s";
        }

        static async Task<T> ExecuteWithRetryAsync<T>(Func<Task<T>> action, string actionName, string token)
        {
            Exception lastException = null;
            for (int attempt = 1; attempt <= GlobalConfig.RetryCount; attempt++)
            {
                try
                {
                    return await action();
                }
                catch (Exception ex) when (attempt < GlobalConfig.RetryCount && IsRetryableException(ex))
                {
                    lastException = ex;
                    var delayMs = 600 * attempt;
                    LogHelper.LogWarn($"{actionName}失败，将进行第{attempt + 1}次重试（总重试次数{GlobalConfig.RetryCount}），Token:{token}，错误：{ex.Message}");
                    await Task.Delay(delayMs);
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    break;
                }
            }

            throw lastException ?? new Exception($"{actionName}失败，Token:{token}");
        }

        static bool IsRetryableException(Exception ex)
        {
            if (ex is HttpRequestException || ex is IOException || ex is TaskCanceledException || ex is TimeoutException)
            {
                return true;
            }

            if (ex.InnerException != null)
            {
                return IsRetryableException(ex.InnerException);
            }

            return false;
        }

        static void UpsertDocumentManifest(ExportDocumentManifestItem item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.DocumentToken))
            {
                return;
            }

            item.UpdatedAtUtc = DateTimeOffset.UtcNow;
            if (!documentManifestByToken.TryGetValue(item.DocumentToken, out var existing))
            {
                documentManifestByToken[item.DocumentToken] = item;
                return;
            }

            existing.NodeToken = string.IsNullOrWhiteSpace(item.NodeToken) ? existing.NodeToken : item.NodeToken;
            existing.ParentDocumentToken = string.IsNullOrWhiteSpace(item.ParentDocumentToken) ? existing.ParentDocumentToken : item.ParentDocumentToken;
            existing.ParentNodeToken = string.IsNullOrWhiteSpace(item.ParentNodeToken) ? existing.ParentNodeToken : item.ParentNodeToken;
            existing.DocumentType = string.IsNullOrWhiteSpace(item.DocumentType) ? existing.DocumentType : item.DocumentType;
            existing.Title = string.IsNullOrWhiteSpace(item.Title) ? existing.Title : item.Title;
            existing.RelativeOutputPath = string.IsNullOrWhiteSpace(item.RelativeOutputPath) ? existing.RelativeOutputPath : item.RelativeOutputPath;
            existing.Status = string.IsNullOrWhiteSpace(item.Status) ? existing.Status : item.Status;
            existing.Error = string.IsNullOrWhiteSpace(item.Error) ? existing.Error : item.Error;
            existing.UpdatedAtUtc = item.UpdatedAtUtc;
        }

        static void UpdateDocumentManifestStatus(string token, string status, string filePath, string error)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return;
            }

            var item = new ExportDocumentManifestItem
            {
                DocumentToken = token,
                Status = status,
                Error = error,
                RelativeOutputPath = string.IsNullOrWhiteSpace(filePath) ? null : ToRelativePath(filePath),
                UpdatedAtUtc = DateTimeOffset.UtcNow
            };

            UpsertDocumentManifest(item);
        }

        static void UpsertAttachmentManifest(ExportAttachmentManifestItem item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.DocumentToken) || string.IsNullOrWhiteSpace(item.FileToken))
            {
                return;
            }

            var key = $"{item.DocumentToken}|{item.BlockId}|{item.FileToken}";
            item.UpdatedAtUtc = DateTimeOffset.UtcNow;
            if (!attachmentManifestByKey.TryGetValue(key, out var existing))
            {
                attachmentManifestByKey[key] = item;
                return;
            }

            if (GetAttachmentStatusPriority(item.Status) >= GetAttachmentStatusPriority(existing.Status))
            {
                attachmentManifestByKey[key] = item;
                return;
            }

            existing.RelativeOutputPath = string.IsNullOrWhiteSpace(existing.RelativeOutputPath) ? item.RelativeOutputPath : existing.RelativeOutputPath;
            existing.DocumentRelativePath = string.IsNullOrWhiteSpace(existing.DocumentRelativePath) ? item.DocumentRelativePath : existing.DocumentRelativePath;
            existing.Sha256 = string.IsNullOrWhiteSpace(existing.Sha256) ? item.Sha256 : existing.Sha256;
            existing.Error = string.IsNullOrWhiteSpace(existing.Error) ? item.Error : existing.Error;
            existing.UpdatedAtUtc = item.UpdatedAtUtc;
        }

        static int GetAttachmentStatusPriority(string status)
        {
            return status switch
            {
                "success" => 4,
                "skipped_existing" => 3,
                "dedup_in_document" => 2,
                "failed" => 1,
                _ => 0
            };
        }

        static async Task WriteManifestFilesAsync()
        {
            var docManifestPath = Path.Combine(GlobalConfig.ExportPath, "export-doc-manifest.json");
            var attachmentManifestPath = Path.Combine(GlobalConfig.ExportPath, "export-attachment-manifest.jsonl");

            var root = new ExportDocumentManifestRoot
            {
                GeneratedAtUtc = DateTimeOffset.UtcNow,
                SourceType = GlobalConfig.Type,
                SaveType = GlobalConfig.DocSaveType,
                WikiSpaceId = GlobalConfig.WikiSpaceId,
                RootToken = string.Equals(GlobalConfig.Type, "wiki", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(GlobalConfig.Type, "wikiToken", StringComparison.OrdinalIgnoreCase)
                    ? GlobalConfig.WikiRootNodeToken
                    : GlobalConfig.CloudDocFolder,
                ExportPath = GlobalConfig.ExportPath,
                Documents = documentManifestByToken.Values
                    .OrderBy(x => x.RelativeOutputPath, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(x => x.DocumentToken, StringComparer.Ordinal)
                    .ToList()
            };

            var jsonOptions = new JsonSerializerOptions
            {
                WriteIndented = true
            };
            await File.WriteAllTextAsync(docManifestPath, JsonSerializer.Serialize(root, jsonOptions));

            var lines = attachmentManifestByKey.Values
                .OrderBy(x => x.DocumentToken, StringComparer.Ordinal)
                .ThenBy(x => x.BlockIndex)
                .ThenBy(x => x.FileToken, StringComparer.Ordinal)
                .Select(x => JsonSerializer.Serialize(x));
            await File.WriteAllLinesAsync(attachmentManifestPath, lines);

            Console.WriteLine($"文档清单已保存：{docManifestPath}");
            Console.WriteLine($"附件清单已保存：{attachmentManifestPath}");
        }

        static string ToRelativePath(string absolutePath)
        {
            var relativePath = Path.GetRelativePath(GlobalConfig.ExportPath, absolutePath);
            return relativePath.Replace("\\", "/");
        }

        static string ComputeSha256(byte[] content)
        {
            using var sha = SHA256.Create();
            var hashBytes = sha.ComputeHash(content ?? Array.Empty<byte>());
            return Convert.ToHexString(hashBytes);
        }

        static string ComputeSha256FromFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return null;
            }

            using var stream = File.OpenRead(filePath);
            using var sha = SHA256.Create();
            var hashBytes = sha.ComputeHash(stream);
            return Convert.ToHexString(hashBytes);
        }

        static string EnsureFilePathWithExtension(string filePath, string sourceFileName)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new Exception("文件保存路径为空，无法保存文件");
            }

            if (!string.IsNullOrWhiteSpace(Path.GetExtension(filePath)))
            {
                return filePath;
            }

            var sourceExt = Path.GetExtension(sourceFileName);
            if (string.IsNullOrWhiteSpace(sourceExt))
            {
                sourceExt = ".bin";
            }

            return filePath + sourceExt;
        }

        static void LoadAttachmentFailureRecords()
        {
            var filePath = GetAttachmentFailureJsonPath();
            if (!File.Exists(filePath))
            {
                return;
            }

            foreach (var line in File.ReadLines(filePath))
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                try
                {
                    var record = JsonSerializer.Deserialize<AttachmentFailureRecord>(line);
                    if (record == null)
                    {
                        continue;
                    }

                    record.FailureKey = string.IsNullOrWhiteSpace(record.FailureKey)
                        ? BuildAttachmentFailureKey(record.DocumentToken, record.BlockId, record.FileToken, record.FileName)
                        : record.FailureKey;
                    if (record.FailureCount <= 0)
                    {
                        record.FailureCount = 1;
                    }
                    if (record.LastFailureAtUtc == default)
                    {
                        record.LastFailureAtUtc = DateTimeOffset.UtcNow;
                    }

                    attachmentFailureByKey[record.FailureKey] = record;
                }
                catch (Exception ex)
                {
                    LogHelper.LogWarn($"解析附件失败记录异常，已忽略。文件：{filePath}，内容：{line}，错误：{ex.Message}");
                }
            }
        }

        static async Task WriteAttachmentFailureFilesAsync()
        {
            var failedRecords = attachmentFailureByKey.Values
                .OrderBy(x => x.DocumentToken, StringComparer.Ordinal)
                .ThenBy(x => x.BlockIndex)
                .ThenBy(x => x.FileToken, StringComparer.Ordinal)
                .ToList();
            var textPath = GetAttachmentFailureTextPath();
            var jsonPath = GetAttachmentFailureJsonPath();

            if (!failedRecords.Any())
            {
                if (File.Exists(textPath))
                {
                    File.Delete(textPath);
                }
                if (File.Exists(jsonPath))
                {
                    File.Delete(jsonPath);
                }
                return;
            }

            Console.WriteLine("以下是附件下载失败清单（已跳过，不影响主流程）");
            for (int i = 0; i < failedRecords.Count; i++)
            {
                Console.WriteLine($"{i + 1}.【{FormatAttachmentFailureMessage(failedRecords[i])}】");
            }

            var textLines = failedRecords.Select(FormatAttachmentFailureMessage);
            var jsonLines = failedRecords.Select(x => JsonSerializer.Serialize(x));
            await File.WriteAllLinesAsync(textPath, textLines);
            await File.WriteAllLinesAsync(jsonPath, jsonLines);
            Console.WriteLine($"附件失败清单已保存：{textPath}");
            Console.WriteLine($"附件失败结构化日志已保存：{jsonPath}");
        }

        static async Task RetryPendingAttachmentFailuresAsync()
        {
            if (!GlobalConfig.Resume || !attachmentFailureByKey.Any())
            {
                return;
            }

            var retryList = attachmentFailureByKey.Values
                .Where(IsRetryableAttachmentFailure)
                .OrderBy(x => x.LastFailureAtUtc)
                .ToList();
            if (!retryList.Any())
            {
                return;
            }

            Console.WriteLine($"开始重试历史附件失败项，共{retryList.Count}条（仅重试可识别fileToken的附件）...");
            var downloadDir = Path.Combine(GlobalConfig.ExportPath, "attachments", "by-token");

            foreach (var record in retryList)
            {
                discoveredAttachmentCount++;
                heartbeatCurrentDocument = record.DocumentToken ?? heartbeatCurrentDocument;
                heartbeatCurrentAttachment = record.FileName ?? heartbeatCurrentAttachment;

                if (exportProgressStore.TryGetDownloadedAttachmentPath(record.FileToken, out var existingPath))
                {
                    UpsertAttachmentManifest(new ExportAttachmentManifestItem
                    {
                        DocumentToken = record.DocumentToken,
                        DocumentRelativePath = record.DocumentRelativePath,
                        BlockId = record.BlockId,
                        BlockIndex = record.BlockIndex,
                        BlockType = record.BlockType,
                        FileToken = record.FileToken,
                        FileName = record.FileName,
                        RelativeOutputPath = ToRelativePath(existingPath),
                        Sha256 = ComputeSha256FromFile(existingPath),
                        Status = "skipped_existing",
                        UpdatedAtUtc = DateTimeOffset.UtcNow
                    });
                    ReportAttachmentProgress(record.DocumentToken, record.FileName, "skipped_existing");
                    attachmentFailureByKey.Remove(record.FailureKey);
                    continue;
                }

                try
                {
                    var attachmentBytes = await ExecuteWithRetryAsync(
                        () => feiShuApiCaller.DownLoadAttachmentBestEffort(record.FileToken),
                        "重试下载附件",
                        record.FileToken);
                    var savePath = BuildAttachmentSavePath(downloadDir, record.FileToken, record.FileName);
                    await savePath.Save(attachmentBytes);
                    exportProgressStore?.MarkAttachmentCompleted(record.FileToken, savePath);
                    UpsertAttachmentManifest(new ExportAttachmentManifestItem
                    {
                        DocumentToken = record.DocumentToken,
                        DocumentRelativePath = record.DocumentRelativePath,
                        BlockId = record.BlockId,
                        BlockIndex = record.BlockIndex,
                        BlockType = record.BlockType,
                        FileToken = record.FileToken,
                        FileName = record.FileName,
                        RelativeOutputPath = ToRelativePath(savePath),
                        Sha256 = ComputeSha256(attachmentBytes),
                        Status = "success",
                        UpdatedAtUtc = DateTimeOffset.UtcNow
                    });
                    ReportAttachmentProgress(record.DocumentToken, record.FileName, "success");
                    attachmentFailureByKey.Remove(record.FailureKey);
                }
                catch (Exception ex)
                {
                    record.Error = ex.Message;
                    record.FailureCount = Math.Max(1, record.FailureCount) + 1;
                    record.LastFailureAtUtc = DateTimeOffset.UtcNow;
                    attachmentFailureByKey[record.FailureKey] = record;
                    UpsertAttachmentManifest(new ExportAttachmentManifestItem
                    {
                        DocumentToken = record.DocumentToken,
                        DocumentRelativePath = record.DocumentRelativePath,
                        BlockId = record.BlockId,
                        BlockIndex = record.BlockIndex,
                        BlockType = record.BlockType,
                        FileToken = record.FileToken,
                        FileName = record.FileName,
                        Status = "failed",
                        Error = ex.Message,
                        UpdatedAtUtc = DateTimeOffset.UtcNow
                    });
                    ReportAttachmentProgress(record.DocumentToken, record.FileName, "failed");
                    LogHelper.LogWarn($"附件重试失败：{FormatAttachmentFailureMessage(record)}");
                }
            }

            Console.WriteLine($"附件失败重试结束，剩余失败 {attachmentFailureByKey.Count} 条。");
        }

        static void RecordAttachmentFailure(
            string documentToken,
            string fileName,
            string fileToken,
            string error,
            string blockId = null,
            int blockIndex = 0,
            int blockType = 0,
            string documentRelativePath = null)
        {
            var failureKey = BuildAttachmentFailureKey(documentToken, blockId, fileToken, fileName);
            if (!attachmentFailureByKey.TryGetValue(failureKey, out var record))
            {
                record = new AttachmentFailureRecord
                {
                    FailureKey = failureKey,
                    DocumentToken = documentToken,
                    DocumentRelativePath = documentRelativePath,
                    BlockId = blockId,
                    BlockIndex = blockIndex,
                    BlockType = blockType,
                    FileToken = fileToken,
                    FileName = fileName,
                    FailureCount = 0
                };
            }

            record.Error = error;
            record.FailureCount = Math.Max(0, record.FailureCount) + 1;
            record.LastFailureAtUtc = DateTimeOffset.UtcNow;
            record.DocumentRelativePath = string.IsNullOrWhiteSpace(record.DocumentRelativePath) ? documentRelativePath : record.DocumentRelativePath;
            record.BlockId = string.IsNullOrWhiteSpace(record.BlockId) ? blockId : record.BlockId;
            record.FileToken = string.IsNullOrWhiteSpace(record.FileToken) ? fileToken : record.FileToken;
            record.FileName = string.IsNullOrWhiteSpace(record.FileName) ? fileName : record.FileName;
            attachmentFailureByKey[failureKey] = record;
            LogHelper.LogWarn($"附件下载失败，已忽略。{FormatAttachmentFailureMessage(record)}");
        }

        static void RemoveAttachmentFailure(string documentToken, string blockId, string fileToken, string fileName)
        {
            var key = BuildAttachmentFailureKey(documentToken, blockId, fileToken, fileName);
            if (attachmentFailureByKey.Remove(key))
            {
                return;
            }

            var prefix = $"{documentToken ?? "N/A"}|{blockId ?? "N/A"}|{fileToken ?? "N/A"}|";
            var fuzzyKeys = attachmentFailureByKey.Keys
                .Where(x => x.StartsWith(prefix, StringComparison.Ordinal))
                .ToList();
            foreach (var fuzzyKey in fuzzyKeys)
            {
                attachmentFailureByKey.Remove(fuzzyKey);
            }
        }

        static bool IsRetryableAttachmentFailure(AttachmentFailureRecord record)
        {
            if (record == null)
            {
                return false;
            }

            return !string.IsNullOrWhiteSpace(record.FileToken)
                && !string.Equals(record.FileToken, "N/A", StringComparison.OrdinalIgnoreCase);
        }

        static string BuildAttachmentFailureKey(string documentToken, string blockId, string fileToken, string fileName)
        {
            return $"{documentToken ?? "N/A"}|{blockId ?? "N/A"}|{fileToken ?? "N/A"}|{fileName ?? "N/A"}";
        }

        static string GetAttachmentFailureTextPath()
        {
            return Path.Combine(GlobalConfig.ExportPath, "attachment-download-failed-list.txt");
        }

        static string GetAttachmentFailureJsonPath()
        {
            return Path.Combine(GlobalConfig.ExportPath, "attachment-download-failed-list.jsonl");
        }

        static string FormatAttachmentFailureMessage(AttachmentFailureRecord record)
        {
            return $"文档Token:{record?.DocumentToken}，文件名:{record?.FileName}，文件Token:{record?.FileToken}，BlockId:{record?.BlockId}，失败次数:{record?.FailureCount}，错误:{record?.Error}";
        }

        static string BuildAttachmentSavePath(string saveDir, string fileToken, string fileName)
        {
            saveDir.CreateIfNotExist();

            var cleanFileName = PathNameHelper.SanitizePathSegment(fileName, "attachment.bin");

            if (string.IsNullOrWhiteSpace(Path.GetExtension(cleanFileName)))
            {
                cleanFileName += ".bin";
            }

            var finalName = $"{fileToken}_{cleanFileName}";
            return Path.Combine(saveDir, finalName);
        }

    }
}
