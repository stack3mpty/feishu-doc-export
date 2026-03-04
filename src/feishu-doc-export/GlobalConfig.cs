using Aspose.Words;
using feishu_doc_export.Helper;
using System.Security.Cryptography;
using System.Text;

namespace feishu_doc_export
{
    public static class GlobalConfig
    {
        public static string AppId { get; set; } 

        public static string AppSecret { get; set; } 

        public static string ExportPath { get; set; } 

        public static string WikiSpaceId { get; set; }

        public static string WikiRootNodeToken { get; set; }

        public static string RootToken { get; set; }

        public static string CloudDocFolder { get; set; } 

        public static bool Quit { get; set; }

        public static bool Resume { get; set; } = true;

        public static bool ResumeSkipAttachmentScan { get; set; } = true;

        public static int RetryCount { get; set; } = 3;

        public static string Type { get; set; } = "wiki";

        public static string Platform { get; set; } = FeiShuConsts.PlatformFeishu;

        public static string AuthMode { get; set; } = "tenant";

        public static string UserAccessToken { get; set; }

        public static string UserRefreshToken { get; set; }

        private static string _docSaveType = "docx";

        public static string DocSaveType { 
            get { return _docSaveType; }
            set
            {
                var options = new string[] { "pdf", "docx", "md" };

                _docSaveType = options.Contains(value) ? value : "docx";
            } 
        }

        /// <summary>
        /// 飞书支持导出的文件类型和导出格式
        /// </summary>
        static Dictionary<string, string> fileExtensionDict = new Dictionary<string, string>()
        {
            {"doc","docx" },
            {"docx","docx" },
            {"sheet","xlsx" },
            {"bitable","xlsx" },
            {"file","file" },
        };

        /// <summary>
        /// 获取飞书支持导出的文件格式
        /// </summary>
        /// <param name="objType"></param>
        /// <param name="fileExt"></param>
        /// <returns></returns>
        public static bool GetFileExtension(string objType, out string fileExt)
        {
            return fileExtensionDict.TryGetValue(objType, out fileExt);
        }

        private static void InitAsposeLicense()
        {
            var license = new License();
            var candidatePaths = new[]
            {
                Environment.GetEnvironmentVariable("ASPOSE_WORDS_LICENSE_PATH"),
                Path.Combine(AppContext.BaseDirectory, "Aspose.lic"),
                Path.Combine(Directory.GetCurrentDirectory(), "Aspose.lic"),
                "C:\\Users\\User\\Desktop\\Aspose.lic"
            }.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct();

            foreach (var path in candidatePaths)
            {
                if (!File.Exists(path))
                {
                    continue;
                }

                try
                {
                    license.SetLicense(path);
                    LogHelper.LogInfo($"已加载Aspose许可：{path}");
                    return;
                }
                catch (Exception ex)
                {
                    LogHelper.LogWarn($"加载Aspose许可失败，路径：{path}，错误：{ex.Message}");
                }
            }

            LogHelper.LogWarn("未找到可用的Aspose许可文件，将以试用模式运行（仅影响文档转换相关能力）。");
        }

        /// <summary>
        /// 初始化全局配置信息
        /// </summary>
        /// <param name="args"></param>
        public static void Init(string[] args)
        {
            if (args.Length > 0)
            {
                AppId = GetCommandLineArg(args, "--appId=");
                AppSecret = GetCommandLineArg(args, "--appSecret=");
                Type = GetCommandLineArg(args, "--type=", true);
                Platform = GetCommandLineArg(args, "--platform=", true);
                AuthMode = GetCommandLineArg(args, "--authMode=", true);
                RootToken = GetCommandLineArg(args, "--rootToken=", true);
                CloudDocFolder = GetCommandLineArg(args, "--folderToken=", true);
                WikiSpaceId = GetCommandLineArg(args, "--spaceId=", true);
                WikiRootNodeToken = GetCommandLineArg(args, "--rootNodeToken=", true);
                DocSaveType = GetCommandLineArg(args, "--saveType=", true);
                ExportPath = GetCommandLineArg(args, "--exportPath=");
                UserAccessToken = GetCommandLineArg(args, "--userAccessToken=", true);
                UserRefreshToken = GetCommandLineArg(args, "--userRefreshToken=", true);
                Quit = args.Contains("--quit");

                var resumeArg = GetCommandLineArg(args, "--resume=", true);
                if (!string.IsNullOrWhiteSpace(resumeArg))
                {
                    if (TryParseBoolArg(resumeArg, out var resume))
                    {
                        Resume = resume;
                    }
                    else
                    {
                        LogHelper.LogWarn($"参数--resume={resumeArg}无效，已使用默认值true。支持值：true/false/1/0/yes/no");
                    }
                }

                var retryCountArg = GetCommandLineArg(args, "--retryCount=", true);
                if (!string.IsNullOrWhiteSpace(retryCountArg))
                {
                    if (int.TryParse(retryCountArg, out var retryCount) && retryCount > 0)
                    {
                        RetryCount = retryCount;
                    }
                    else
                    {
                        LogHelper.LogWarn($"参数--retryCount={retryCountArg}无效，已使用默认值3。");
                    }
                }

                var resumeSkipAttachmentScanArg = GetCommandLineArg(args, "--resumeSkipAttachmentScan=", true);
                if (!string.IsNullOrWhiteSpace(resumeSkipAttachmentScanArg))
                {
                    if (TryParseBoolArg(resumeSkipAttachmentScanArg, out var resumeSkipAttachmentScan))
                    {
                        ResumeSkipAttachmentScan = resumeSkipAttachmentScan;
                    }
                    else
                    {
                        LogHelper.LogWarn($"参数--resumeSkipAttachmentScan={resumeSkipAttachmentScanArg}无效，已使用默认值true。支持值：true/false/1/0/yes/no");
                    }
                }
            }
            else
            {
//#if !DEBUG
                Console.WriteLine("请输入飞书自建应用的AppId：");
                AppId = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(AppId))
                {
                    LogHelper.LogWarnExit("AppId是必填参数");
                }

                Console.WriteLine("请输入飞书自建应用的AppSecret：");
                AppSecret = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(AppSecret))
                {
                    LogHelper.LogWarnExit("AppSecret是必填参数");
                }

                Console.WriteLine("请输入文档导出的文件类型（可选值：docx、md、pdf，为空或其他非可选值则默认为docx）：");
                DocSaveType = Console.ReadLine();

                Console.WriteLine("请选择云文档类型（可选值：wiki、wikiToken、cloudDoc）");
                Type = Console.ReadLine();

                Console.WriteLine("请选择平台（可选值：feishu、lark；为空默认feishu）");
                Platform = Console.ReadLine();

                Console.WriteLine("请选择鉴权模式（可选值：tenant、user；为空默认tenant）");
                AuthMode = Console.ReadLine();
                if (string.Equals(AuthMode, "user", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("请输入用户user_access_token（必填）");
                    UserAccessToken = Console.ReadLine();
                }

                if (string.Equals(Type, "cloudDoc", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("请输入起始节点Token（支持 --rootToken 语义；cloudDoc 模式下可填文件夹或父文档token，必填项！）");
                    CloudDocFolder = Console.ReadLine();
                    if (string.IsNullOrWhiteSpace(CloudDocFolder))
                    {
                        LogHelper.LogWarnExit("起始节点Token是必填参数");
                    }
                }
                else
                {
                    if (!string.Equals(Type, "wikiToken", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine("请输入要导出的知识库Id（为空代表从所有知识库中选择）：");
                        WikiSpaceId = Console.ReadLine();
                    }

                    Console.WriteLine("请输入要导出的父文档NodeToken（支持 --rootToken 语义；wikiToken模式必填，wiki模式为空代表导出知识库全部文档）：");
                    WikiRootNodeToken = Console.ReadLine();
                }

                Console.WriteLine("请输入文档导出的目录位置：");
                ExportPath = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(ExportPath))
                {
                    LogHelper.LogWarnExit("文档导出的目录是必填参数");
                }
                //#endif
            }

            NormalizeArgs();
            InitAsposeLicense();
        }

        private static void NormalizeArgs()
        {
            Type = (Type ?? "wiki").Trim();
            if (!string.Equals(Type, "wiki", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(Type, "wikiToken", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(Type, "cloudDoc", StringComparison.OrdinalIgnoreCase))
            {
                LogHelper.LogWarn($"参数--type={Type}不在支持范围内，已自动使用wiki模式。");
                Type = "wiki";
            }

            Platform = (Platform ?? FeiShuConsts.PlatformFeishu).Trim().ToLowerInvariant();
            if (string.Equals(Platform, "openfeishu", StringComparison.OrdinalIgnoreCase))
            {
                Platform = FeiShuConsts.PlatformFeishu;
            }
            if (!string.Equals(Platform, FeiShuConsts.PlatformFeishu, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(Platform, FeiShuConsts.PlatformLark, StringComparison.OrdinalIgnoreCase))
            {
                LogHelper.LogWarn($"参数--platform={Platform}不在支持范围内，已自动使用feishu。");
                Platform = FeiShuConsts.PlatformFeishu;
            }

            AuthMode = (AuthMode ?? "tenant").Trim().ToLowerInvariant();
            if (!string.Equals(AuthMode, "tenant", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(AuthMode, "user", StringComparison.OrdinalIgnoreCase))
            {
                LogHelper.LogWarn($"参数--authMode={AuthMode}不在支持范围内，已自动使用tenant。");
                AuthMode = "tenant";
            }
            if (string.Equals(AuthMode, "user", StringComparison.OrdinalIgnoreCase) &&
                string.IsNullOrWhiteSpace(UserAccessToken))
            {
                LogHelper.LogWarnExit("user鉴权模式下，--userAccessToken是必填参数。");
            }

            if (string.Equals(Type, "wiki", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(Type, "wikiToken", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrWhiteSpace(WikiRootNodeToken))
                {
                    if (!string.IsNullOrWhiteSpace(RootToken))
                    {
                        WikiRootNodeToken = RootToken;
                    }
                    else if (!string.IsNullOrWhiteSpace(CloudDocFolder))
                    {
                        WikiRootNodeToken = CloudDocFolder;
                        LogHelper.LogWarn("检测到wiki模式下传入了--folderToken，已自动将其视为--rootNodeToken。");
                    }
                }

                if (!string.IsNullOrWhiteSpace(RootToken) &&
                    !string.IsNullOrWhiteSpace(WikiRootNodeToken) &&
                    !string.Equals(RootToken, WikiRootNodeToken, StringComparison.Ordinal))
                {
                    LogHelper.LogWarn("参数--rootToken与--rootNodeToken不一致，优先使用--rootNodeToken。");
                }

                if (string.Equals(Type, "wikiToken", StringComparison.OrdinalIgnoreCase) &&
                    string.IsNullOrWhiteSpace(WikiRootNodeToken))
                {
                    LogHelper.LogWarnExit("wikiToken模式下，--rootNodeToken（或--rootToken）是必填参数。");
                }
            }
            else
            {
                if (string.IsNullOrWhiteSpace(CloudDocFolder))
                {
                    if (!string.IsNullOrWhiteSpace(RootToken))
                    {
                        CloudDocFolder = RootToken;
                    }
                    else if (!string.IsNullOrWhiteSpace(WikiRootNodeToken))
                    {
                        CloudDocFolder = WikiRootNodeToken;
                        LogHelper.LogWarn("检测到cloudDoc模式下传入了--rootNodeToken，已自动将其视为--folderToken。");
                    }
                }

                if (!string.IsNullOrWhiteSpace(RootToken) &&
                    !string.IsNullOrWhiteSpace(CloudDocFolder) &&
                    !string.Equals(RootToken, CloudDocFolder, StringComparison.Ordinal))
                {
                    LogHelper.LogWarn("参数--rootToken与--folderToken不一致，优先使用--folderToken。");
                }
            }
        }

        /// <summary>
        /// 获取命令行参数值
        /// </summary>
        /// <param name="args"></param>
        /// <param name="parameterName"></param>
        /// <returns></returns>
        public static string GetCommandLineArg(string[] args, string parameterName, bool canNull = false)
        {
            // 参数值
            string paraValue = string.Empty;
            // 是否有匹配的参数
            bool found = false;
            foreach (string arg in args)
            {
                if (arg.StartsWith(parameterName))
                {
                    paraValue = arg.Substring(parameterName.Length);
                    found = true;
                }
            }

            if (!canNull)
            {
                if (!found)
                {
                    Console.WriteLine($"没有找到参数：{parameterName}");
                    Console.WriteLine("请填写以下所有参数：");
                    Console.WriteLine("  --appId           飞书自建应用的AppId.【必填项】");
                    Console.WriteLine("  --appSecret       飞书自建应用的AppSecret.【必填项】");
                    Console.WriteLine("  --exportPath      文档导出的目录位置.【必填项】");
                    Console.WriteLine("  --type            知识库（wiki/wikiToken）或个人空间云文档（cloudDoc）（可选值：wiki、wikiToken、cloudDoc，为空则默认为wiki）");
                    Console.WriteLine("  --platform        平台环境（可选值：feishu、lark，默认feishu）");
                    Console.WriteLine("  --authMode        鉴权模式（可选值：tenant、user，默认tenant）");
                    Console.WriteLine("  --userAccessToken user模式下的用户访问令牌（authMode=user时必填）");
                    Console.WriteLine("  --userRefreshToken user模式下的用户刷新令牌（预留，可选）");
                    Console.WriteLine("  --saveType        文档导出的文件类型（可选值：docx、md、pdf，为空或其他非可选值则默认为docx）.");
                    Console.WriteLine("  --rootToken       通用起始节点token（wiki/wikiToken=父文档node_token，cloudDoc=文件夹/父文档token）.");
                    Console.WriteLine("  --folderToken     当type为个人空间云文档时，该项必填（支持文件夹或父文档token）");
                    Console.WriteLine("  --spaceId         飞书导出的知识库Id.");
                    Console.WriteLine("  --rootNodeToken   当type为wiki/wikiToken时，指定要导出的父文档节点（导出该节点及全部后代）。wikiToken模式下必填。");
                    Console.WriteLine("  --resume          是否启用断点续传（可选值：true/false/1/0，默认true）.");
                    Console.WriteLine("  --resumeSkipAttachmentScan  断点续传时是否跳过已完成文档的附件块扫描（可选值：true/false/1/0，默认true）.");
                    Console.WriteLine("  --retryCount      下载重试次数（默认3）.");
                    Environment.Exit(0);
                }

                // 参数值为空
                if (string.IsNullOrWhiteSpace(paraValue))
                {
                    Console.WriteLine($"参数{parameterName}不能为空");
                    Environment.Exit(0);
                }
            }

            return paraValue;
        }

        private static bool TryParseBoolArg(string value, out bool result)
        {
            result = false;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var normalized = value.Trim().ToLowerInvariant();
            if (normalized == "true" || normalized == "1" || normalized == "yes" || normalized == "y")
            {
                result = true;
                return true;
            }

            if (normalized == "false" || normalized == "0" || normalized == "no" || normalized == "n")
            {
                result = false;
                return true;
            }

            return false;
        }
    }
}
