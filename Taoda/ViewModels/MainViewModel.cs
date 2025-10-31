using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Taoda.Models;
using Taoda.Services;

namespace Taoda.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    // 服务依赖
    private readonly IExcelService _excelService;
    private readonly IWordService _wordService;
    private readonly IFileDialogService _fileDialogService;

    // 可观察属性
    [ObservableProperty] private string? _excelFilePath;
    [ObservableProperty] private string? _templateFilePath;
    [ObservableProperty] private string? _outputDirectory;
    [ObservableProperty] private ExcelDataModel? _excelData;
    [ObservableProperty] private TemplateModel? _templateModel;
    [ObservableProperty] private VariableMatchingModel? _variableMatching;
    [ObservableProperty] private string _statusMessage = "就绪";
    [ObservableProperty] private bool _isProcessing;
    [ObservableProperty] private int _progressValue;
    [ObservableProperty] private int _progressMaximum = 100;
    [ObservableProperty] private bool _hasError;
    [ObservableProperty] private string? _errorMessage;

    // 命令
    public IAsyncRelayCommand SelectExcelFileCommand { get; }
    public IAsyncRelayCommand SelectTemplateFileCommand { get; }
    public IAsyncRelayCommand SelectOutputDirectoryCommand { get; }
    public IAsyncRelayCommand GenerateDocumentsCommand { get; }

    // 计算属性
    public bool CanGenerate => !string.IsNullOrEmpty(ExcelFilePath) &&
                               !string.IsNullOrEmpty(TemplateFilePath) &&
                               !string.IsNullOrEmpty(OutputDirectory) &&
                               !IsProcessing;

    public MainViewModel(IExcelService excelService, IWordService wordService, IFileDialogService fileDialogService)
    {
        _excelService = excelService;
        _wordService = wordService;
        _fileDialogService = fileDialogService;

        // 初始化命令
        SelectExcelFileCommand = new AsyncRelayCommand(SelectExcelFileAsync);
        SelectTemplateFileCommand = new AsyncRelayCommand(SelectTemplateFileAsync);
        SelectOutputDirectoryCommand = new AsyncRelayCommand(SelectOutputDirectoryAsync);
        GenerateDocumentsCommand = new AsyncRelayCommand(GenerateDocumentsAsync, () => CanGenerate);

        // 设置属性变更通知，当相关属性变化时更新命令可执行状态
        PropertyChanged += (_, e) =>
        {
            if (e.PropertyName is nameof(ExcelFilePath) or nameof(TemplateFilePath) or nameof(OutputDirectory)
                or nameof(IsProcessing))
            {
                GenerateDocumentsCommand.NotifyCanExecuteChanged();
                OnPropertyChanged(nameof(CanGenerate));
            }
        };
    }

    // 命令实现方法
    private async Task SelectExcelFileAsync()
    {
        try
        {
            StatusMessage = "选择Excel文件...";

            var filters = new[]
            {
                new FileDialogFilter { Name = "Excel文件", Extensions = ["xlsx", "xls"] },
                new FileDialogFilter { Name = "所有文件", Extensions = ["*"] }
            };

            var filePath = await _fileDialogService.OpenFileAsync("选择Excel数据文件", filters);

            if (!string.IsNullOrEmpty(filePath))
            {
                // 验证文件格式
                var extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls")
                {
                    StatusMessage = "错误：请选择有效的Excel文件格式（.xlsx 或 .xls）";
                    return;
                }

                ExcelFilePath = filePath;
                UpdateStatus($"已选择Excel文件：{Path.GetFileName(filePath)}");

                // 自动加载Excel数据
                await LoadExcelDataAsync();
            }
            else
            {
                StatusMessage = "未选择文件";
            }
        }
        catch (Exception ex)
        {
            HandleError(ex, "选择Excel文件时发生错误");
        }
    }

    private async Task SelectTemplateFileAsync()
    {
        try
        {
            StatusMessage = "选择Word模板...";

            var filters = new[]
            {
                new FileDialogFilter { Name = "Word文档", Extensions = ["docx"] },
                new FileDialogFilter { Name = "所有文件", Extensions = ["*"] }
            };

            var filePath = await _fileDialogService.OpenFileAsync("选择Word模板文件", filters);

            if (!string.IsNullOrEmpty(filePath))
            {
                // 验证文件格式
                var extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".docx")
                {
                    StatusMessage = "错误：请选择有效的Word文档格式（.docx）";
                    return;
                }

                TemplateFilePath = filePath;
                UpdateStatus($"已选择模板文件：{Path.GetFileName(filePath)}");

                // 自动加载模板变量
                await LoadTemplateVariablesAsync();
            }
            else
            {
                StatusMessage = "未选择文件";
            }
        }
        catch (Exception ex)
        {
            HandleError(ex, "选择Word模板时发生错误");
        }
    }

    private async Task SelectOutputDirectoryAsync()
    {
        try
        {
            StatusMessage = "选择输出目录...";

            var directoryPath = await _fileDialogService.SelectFolderAsync("选择文档输出目录");

            if (!string.IsNullOrEmpty(directoryPath))
            {
                // 验证目录写入权限
                if (!Directory.Exists(directoryPath))
                {
                    StatusMessage = "错误：所选目录不存在";
                    return;
                }

                try
                {
                    // 测试写入权限
                    var testFile = Path.Combine(directoryPath, $"test_{Guid.NewGuid()}.tmp");
                    await File.WriteAllTextAsync(testFile, "test");
                    File.Delete(testFile);

                    OutputDirectory = directoryPath;
                    UpdateStatus($"已选择输出目录：{directoryPath}");
                }
                catch (UnauthorizedAccessException)
                {
                    StatusMessage = "错误：对所选目录没有写入权限";
                }
                catch (Exception ex)
                {
                    HandleError(ex, "验证目录权限时发生错误");
                }
            }
            else
            {
                StatusMessage = "未选择目录";
            }
        }
        catch (Exception ex)
        {
            HandleError(ex, "选择输出目录时发生错误");
        }
    }

    private async Task GenerateDocumentsAsync()
    {
        if (ExcelData == null || TemplateModel == null || string.IsNullOrEmpty(OutputDirectory))
        {
            StatusMessage = "错误：请确保已选择Excel文件、Word模板和输出目录";
            return;
        }

        try
        {
            IsProcessing = true;
            ProgressValue = 0;
            ProgressMaximum = ExcelData.TotalRows;
            StatusMessage = "开始生成文档...";

            var progress = new Progress<int>(value =>
            {
                ProgressValue = value;
                StatusMessage = $"正在生成文档... ({value}/{ProgressMaximum})";
            });

            // 执行批量文档生成
            var generatedCount = await _wordService.BatchGenerateAsync(
                TemplateModel.FilePath,
                ExcelData.Rows,
                OutputDirectory,
                progress);

            // 生成完成
            ProgressValue = ProgressMaximum;
            UpdateStatus($"文档生成完成！成功生成 {generatedCount} 个文档，保存在：{OutputDirectory}");

            // 可选：打开输出目录
            // Process.Start("explorer.exe", OutputDirectory);
        }
        catch (Exception ex)
        {
            HandleError(ex, "文档生成失败");
        }
        finally
        {
            IsProcessing = false;
        }
    }

    // 数据加载和处理方法
    private async Task LoadExcelDataAsync()
    {
        if (string.IsNullOrEmpty(ExcelFilePath))
            return;

        try
        {
            StatusMessage = "正在读取Excel文件...";
            IsProcessing = true;

            // 读取Excel数据
            ExcelData = await _excelService.ReadExcelAsync(ExcelFilePath);

            UpdateStatus($"Excel数据加载完成：检测到 {ExcelData.Headers.Count} 个变量，{ExcelData.TotalRows} 行数据");

            // 如果模板已选择，执行变量匹配
            if (TemplateModel != null)
            {
                await PerformVariableMatchingAsync();
            }
        }
        catch (Exception ex)
        {
            HandleError(ex, "读取Excel文件失败");
            ExcelData = null;
        }
        finally
        {
            IsProcessing = false;
        }
    }

    private async Task LoadTemplateVariablesAsync()
    {
        if (string.IsNullOrEmpty(TemplateFilePath))
            return;

        try
        {
            StatusMessage = "正在扫描模板变量...";
            IsProcessing = true;

            // 提取模板变量
            var variables = _wordService.ExtractTemplateVariables(TemplateFilePath);

            TemplateModel = new TemplateModel
            {
                FilePath = TemplateFilePath!,
                Variables = variables
            };

            UpdateStatus($"模板变量扫描完成：检测到 {variables.Count} 个变量");

            // 如果Excel数据已加载，执行变量匹配
            if (ExcelData != null)
            {
                await PerformVariableMatchingAsync();
            }
        }
        catch (Exception ex)
        {
            HandleError(ex, "扫描模板变量失败");
            TemplateModel = null;
        }
        finally
        {
            IsProcessing = false;
        }
    }

    private async Task PerformVariableMatchingAsync()
    {
        if (ExcelData == null || TemplateModel == null)
            return;

        try
        {
            StatusMessage = "正在进行变量匹配...";

            await Task.Run(() =>
            {
                var matching = new VariableMatchingModel();

                // 精确匹配
                var exactMatches = ExcelData.Headers.Intersect(TemplateModel.Variables).ToList();
                matching.MatchedVariables.AddRange(exactMatches);

                // 未匹配的项目
                matching.UnmatchedExcelColumns = ExcelData.Headers.Except(exactMatches).ToList();
                matching.UnmatchedTemplateVariables = TemplateModel.Variables.Except(exactMatches).ToList();

                VariableMatching = matching;
            });

            if (VariableMatching == null) return;

            var matchedCount = VariableMatching.MatchedVariables.Count;

            if (VariableMatching.IsFullyMatched)
            {
                UpdateStatus($"变量匹配完成：所有 {matchedCount} 个变量完全匹配");
            }
            else
            {
                UpdateStatus(
                    $"变量匹配完成：{matchedCount} 个匹配，{VariableMatching.UnmatchedExcelColumns.Count} 个Excel变量未匹配，{VariableMatching.UnmatchedTemplateVariables.Count} 个模板变量未匹配");
            }
        }
        catch (Exception ex)
        {
            HandleError(ex, "变量匹配失败");
            VariableMatching = null;
        }
    }

    // 状态管理和错误处理方法
    private void HandleError(Exception ex, string context)
    {
        HasError = true;

        var errorMessage = ex switch
        {
            FileNotFoundException => $"文件未找到：{ex.Message}",
            UnauthorizedAccessException => $"文件访问权限不足：{ex.Message}",
            DirectoryNotFoundException => $"目录未找到：{ex.Message}",
            IOException => $"文件操作错误：{ex.Message}",
            InvalidOperationException => $"操作无效：{ex.Message}",
            ArgumentException => $"参数错误：{ex.Message}",
            _ => $"未知错误：{ex.Message}"
        };

        ErrorMessage = errorMessage;
        StatusMessage = $"{context}：{errorMessage}";

        // 记录详细错误信息（可以扩展为日志记录）
        System.Diagnostics.Debug.WriteLine($"[ERROR] {context}: {ex}");
    }

    private void ClearError()
    {
        HasError = false;
        ErrorMessage = null;
    }

    private void UpdateStatus(string message, bool isError = false)
    {
        StatusMessage = message;
        if (!isError)
        {
            ClearError();
        }
    }

    private void ResetProgress()
    {
        ProgressValue = 0;
        ProgressMaximum = 100;
    }

    // 重写属性变更方法以增强状态管理
    partial void OnIsProcessingChanged(bool value)
    {
        if (!value)
        {
            // 处理完成时重置进度条
            if (ProgressValue >= ProgressMaximum && ProgressMaximum > 0)
            {
                // 保持完成状态一段时间，然后重置
                Task.Delay(2000).ContinueWith(_ =>
                {
                    if (!IsProcessing)
                    {
                        ResetProgress();
                    }
                });
            }
        }

        // 通知相关命令更新可执行状态
        GenerateDocumentsCommand.NotifyCanExecuteChanged();
        OnPropertyChanged(nameof(CanGenerate));
    }

    partial void OnExcelFilePathChanged(string? value)
    {
        if (!string.IsNullOrEmpty(value))
        {
            ClearError();
        }
    }

    partial void OnTemplateFilePathChanged(string? value)
    {
        if (!string.IsNullOrEmpty(value))
        {
            ClearError();
        }
    }

    partial void OnOutputDirectoryChanged(string? value)
    {
        if (!string.IsNullOrEmpty(value))
        {
            ClearError();
        }
    }
}
