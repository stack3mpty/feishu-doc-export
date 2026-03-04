using System.Text.Json;

namespace feishu_doc_export.Helper
{
    public class ExportProgressState
    {
        public Dictionary<string, string> CompletedDocuments { get; set; } = new Dictionary<string, string>();

        public Dictionary<string, string> CompletedAttachments { get; set; } = new Dictionary<string, string>();
    }

    public class ExportProgressStore
    {
        private readonly object _syncRoot = new object();
        private readonly string _exportRoot;
        private readonly string _statePath;
        private ExportProgressState _state = new ExportProgressState();

        public string StatePath => _statePath;

        public ExportProgressStore(string exportRoot)
        {
            _exportRoot = Path.GetFullPath(exportRoot ?? ".");
            _statePath = Path.Combine(_exportRoot, ".feishu-export-progress.json");

            Load();
        }

        public bool ShouldSkipDocument(string documentToken, string outputPath)
        {
            if (string.IsNullOrWhiteSpace(documentToken))
            {
                return false;
            }

            if (IsValidFile(outputPath))
            {
                MarkDocumentCompleted(documentToken, outputPath);
                return true;
            }

            lock (_syncRoot)
            {
                if (_state.CompletedDocuments.TryGetValue(documentToken, out var relativePath))
                {
                    var absPath = ToAbsolutePath(relativePath);
                    if (IsValidFile(absPath))
                    {
                        return true;
                    }

                    _state.CompletedDocuments.Remove(documentToken);
                    SaveLocked();
                }
            }

            return false;
        }

        public void MarkDocumentCompleted(string documentToken, string outputPath)
        {
            if (string.IsNullOrWhiteSpace(documentToken) || !IsValidFile(outputPath))
            {
                return;
            }

            lock (_syncRoot)
            {
                _state.CompletedDocuments[documentToken] = ToRelativePath(outputPath);
                SaveLocked();
            }
        }

        public bool TryGetDownloadedAttachmentPath(string attachmentToken, out string filePath)
        {
            filePath = null;
            if (string.IsNullOrWhiteSpace(attachmentToken))
            {
                return false;
            }

            lock (_syncRoot)
            {
                if (!_state.CompletedAttachments.TryGetValue(attachmentToken, out var relativePath))
                {
                    return false;
                }

                var absPath = ToAbsolutePath(relativePath);
                if (IsValidFile(absPath))
                {
                    filePath = absPath;
                    return true;
                }

                _state.CompletedAttachments.Remove(attachmentToken);
                SaveLocked();
                return false;
            }
        }

        public void MarkAttachmentCompleted(string attachmentToken, string outputPath)
        {
            if (string.IsNullOrWhiteSpace(attachmentToken) || !IsValidFile(outputPath))
            {
                return;
            }

            lock (_syncRoot)
            {
                _state.CompletedAttachments[attachmentToken] = ToRelativePath(outputPath);
                SaveLocked();
            }
        }

        private void Load()
        {
            if (!File.Exists(_statePath))
            {
                _state = new ExportProgressState();
                return;
            }

            try
            {
                var json = File.ReadAllText(_statePath);
                _state = JsonSerializer.Deserialize<ExportProgressState>(json) ?? new ExportProgressState();
            }
            catch (Exception ex)
            {
                _state = new ExportProgressState();
                LogHelper.LogWarn($"Failed to load resume state file, start with empty state. File: {_statePath}, Error: {ex.Message}");
            }
        }

        private void SaveLocked()
        {
            try
            {
                var dir = Path.GetDirectoryName(_statePath);
                dir.CreateIfNotExist();

                var tempPath = _statePath + ".tmp";
                var json = JsonSerializer.Serialize(_state, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                File.WriteAllText(tempPath, json);
                File.Move(tempPath, _statePath, true);
            }
            catch (Exception ex)
            {
                LogHelper.LogWarn($"Failed to save resume state file. File: {_statePath}, Error: {ex.Message}");
            }
        }

        private string ToRelativePath(string absPath)
        {
            var fullPath = Path.GetFullPath(absPath);
            return Path.GetRelativePath(_exportRoot, fullPath);
        }

        private string ToAbsolutePath(string relativePath)
        {
            return Path.GetFullPath(Path.Combine(_exportRoot, relativePath ?? string.Empty));
        }

        private static bool IsValidFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            return File.Exists(path) && new FileInfo(path).Length > 0;
        }
    }
}
