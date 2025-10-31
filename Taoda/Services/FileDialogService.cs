using System;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform.Storage;

namespace Taoda.Services;

/// <summary>
/// 文件对话框服务实现类
/// </summary>
public class FileDialogService : IFileDialogService
{
    /// <summary>
    /// 异步打开文件选择对话框
    /// </summary>
    /// <param name="title">对话框标题</param>
    /// <param name="filters">文件过滤器</param>
    /// <returns>选择的文件路径，如果取消则返回 null</returns>
    public async Task<string?> OpenFileAsync(string title, FileDialogFilter[] filters)
    {
        try
        {
            var topLevel = GetTopLevel();
            if (topLevel == null || topLevel.StorageProvider is not { } storageProvider)
                return null;

            // 转换自定义过滤器为 Avalonia 过滤器
            var avaloniaFilters = filters.Select(filter => new FilePickerFileType(filter.Name)
            {
                Patterns = filter.Extensions.Select(ext => ext.StartsWith("*.") ? ext : $"*.{ext}").ToArray()
            }).ToArray();

            var options = new FilePickerOpenOptions
            {
                Title = title,
                AllowMultiple = false,
                FileTypeFilter = avaloniaFilters
            };

            var result = await storageProvider.OpenFilePickerAsync(options);
            var selectedFile = result.FirstOrDefault();
            return selectedFile == null ? null : selectedFile.Path.LocalPath;
        }
        catch (Exception)
        {
            // 如果对话框操作失败，返回 null
            return null;
        }
    }

    /// <summary>
    /// 异步选择文件夹对话框
    /// </summary>
    /// <param name="title">对话框标题</param>
    /// <returns>选择的文件夹路径，如果取消则返回 null</returns>
    public async Task<string?> SelectFolderAsync(string title)
    {
        try
        {
            var topLevel = GetTopLevel();
            if (topLevel == null || topLevel.StorageProvider is not { } storageProvider)
                return null;

            var options = new FolderPickerOpenOptions
            {
                Title = title,
                AllowMultiple = false
            };

            var result = await storageProvider.OpenFolderPickerAsync(options);
            var selectedFolder = result.FirstOrDefault();
            return selectedFolder == null ? null : selectedFolder.Path.LocalPath;
        }
        catch (Exception)
        {
            // 如果对话框操作失败，返回 null
            return null;
        }
    }

    /// <summary>
    /// 获取顶级窗口
    /// </summary>
    /// <returns>顶级窗口实例</returns>
    private static TopLevel? GetTopLevel()
    {
        if (Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            return desktop.MainWindow;
        }

        if (Application.Current?.ApplicationLifetime is ISingleViewApplicationLifetime singleView)
        {
            return TopLevel.GetTopLevel(singleView.MainView);
        }

        return null;
    }
}