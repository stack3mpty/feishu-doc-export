# feishu-doc-export

将 Feishu/Lark 文档导出到本地，支持保留目录结构、正文与附件，并支持断点续传。

## 功能
- 导出父文档及全部后代文档（递归）
- 保留原始目录结构
- 支持导出 `docx` / `md` / `pdf`
- 附件下载到本地（失败不阻断主流程）
- 断点续传与快速续跑
- 导出进度日志与定时心跳
- 支持 `feishu`（国内）与 `lark`（海外）平台
- 支持 `tenant` 与 `user` 两种鉴权模式

## 环境要求
- .NET 6 SDK（构建）
- Windows 可直接运行 `exe`；Linux/macOS 可按对应 RID 发布

## 构建
在仓库根目录执行：

```powershell
dotnet restore src/feishu-doc-export.sln
dotnet build src/feishu-doc-export.sln -c Release
```

Windows 可执行文件位置：
- `src\feishu-doc-export\bin\Release\net6.0\feishu-doc-export.exe`

## 常用参数

必填：
- `--appId=<APP_ID>`
- `--appSecret=<APP_SECRET>`
- `--exportPath=<EXPORT_DIR>`

常用：
- `--platform=feishu|lark`（默认 `feishu`）
- `--type=wiki|wikiToken|cloudDoc`（默认 `wiki`）
- `--saveType=docx|md|pdf`（默认 `docx`）
- `--authMode=tenant|user`（默认 `tenant`）
- `--resume=true|false`（默认 `true`）
- `--resumeSkipAttachmentScan=true|false`（默认 `true`）
- `--retryCount=<N>`（默认 `3`）
- `--quit`（结束后退出）

按模式额外参数：
- `wiki`：`--spaceId=<WIKI_SPACE_ID>`，可选 `--rootNodeToken=<ROOT_NODE_TOKEN>`
- `wikiToken`：`--rootNodeToken=<ROOT_NODE_TOKEN>`（必填）
- `cloudDoc`：`--folderToken=<FOLDER_OR_PARENT_DOC_TOKEN>`（必填）
- `authMode=user`：`--userAccessToken=<USER_ACCESS_TOKEN>`（必填）

## 使用示例

### 1) Lark + wikiToken（推荐用于“指定父文档及其全部后代”）
```powershell
.\feishu-doc-export.exe `
  --appId=cli_xxx `
  --appSecret=xxx `
  --platform=lark `
  --type=wikiToken `
  --rootNodeToken=xxxxxxxx `
  --saveType=md `
  --exportPath=<导出目录> `
  --resume=true `
  --retryCount=3 `
  --quit
```

### 2) Feishu + wiki（按知识库导出）
```powershell
.\feishu-doc-export.exe `
  --appId=cli_xxx `
  --appSecret=xxx `
  --platform=feishu `
  --type=wiki `
  --spaceId=xxxxxxxx `
  --rootNodeToken=xxxxxxxx `
  --saveType=md `
  --exportPath=<导出目录> `
  --resume=true `
  --quit
```

### 3) cloudDoc（Drive 文件夹/父文档）
```powershell
.\feishu-doc-export.exe `
  --appId=cli_xxx `
  --appSecret=xxx `
  --platform=feishu `
  --type=cloudDoc `
  --folderToken=xxxxxxxx `
  --saveType=md `
  --exportPath=<导出目录> `
  --resume=true `
  --quit
```

### 4) user 鉴权（处理权限更细粒度场景）
```powershell
.\feishu-doc-export.exe `
  --appId=cli_xxx `
  --appSecret=xxx `
  --platform=lark `
  --authMode=user `
  --userAccessToken=u-xxx `
  --type=wikiToken `
  --rootNodeToken=xxxxxxxx `
  --saveType=md `
  --exportPath=<导出目录> `
  --resume=true `
  --resumeSkipAttachmentScan=true `
  --quit
```

## 续跑建议
- 日常重跑建议：`--resume=true --resumeSkipAttachmentScan=true`（速度快）
- 需要补齐附件清单时：`--resume=true --resumeSkipAttachmentScan=false`

## 导出产物
导出目录中会生成：
- `export-doc-manifest.json`（文档清单）
- `export-attachment-manifest.jsonl`（附件清单）
- `.feishu-export-progress.json`（断点状态）
- `attachments/by-token/*`（附件实体）
- `attachment-download-failed-list.txt`
- `attachment-download-failed-list.jsonl`
- 正文文件（`md/docx/pdf`）

## 安全说明
- 文档中的 `APP_ID`、`APP_SECRET`、`TOKEN`、`PAGE_ID` 均为占位符示例
- 不要在代码仓库提交真实凭证与导出数据

## 致谢与上游支持
- 本仓库基于上游项目 [eternalfree/feishu-doc-export](https://github.com/eternalfree/feishu-doc-export) 二次开发
- 感谢上游作者与贡献者的开源工作
- 如果你需要更贴近原始实现的版本，请优先参考上游仓库

## 本仓库增强特性
- 增强工程化能力：清晰的清单产物、失败清单、断点状态与可观测日志
- 平台能力增强：同时支持 `feishu` 与 `lark`（通过 `--platform` 切换）
- 内容导出增强：支持导出文档正文，并支持附件与图片落盘
- 稳定性增强：支持断点续传、重试策略、附件失败不阻断主流程

## 许可证与合规说明
- 本仓库沿用并遵守 Apache License 2.0（见仓库根目录 `LICENSE`）
- 保留上游项目的版权与许可证声明，不移除已有 License 头与声明
- 对二次开发部分明确标注为修改/新增内容
- 若上游存在 `NOTICE` 文件，分发时需一并保留并在必要处补充你的新增声明
