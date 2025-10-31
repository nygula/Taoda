using System.Threading.Tasks;

namespace Taoda.Services;

/// <summary>
/// 文件对话框过滤器
/// </summary>
public class FileDialogFilter
{
    /// <summary>
    /// 过滤器名称
    /// </summary>
    public string Name { get; set; } = string.Empty;
    
    /// <summary>
    /// 文件扩展名列表
    /// </summary>
    public string[] Extensions { get; set; } = [];
}

/// <summary>
/// 文件对话框服务接口
/// </summary>
public interface IFileDialogService
{
    /// <summary>
    /// 异步打开文件选择对话框
    /// </summary>
    /// <param name="title">对话框标题</param>
    /// <param name="filters">文件过滤器</param>
    /// <returns>选择的文件路径，如果取消则返回 null</returns>
    Task<string?> OpenFileAsync(string title, FileDialogFilter[] filters);
    
    /// <summary>
    /// 异步选择文件夹对话框
    /// </summary>
    /// <param name="title">对话框标题</param>
    /// <returns>选择的文件夹路径，如果取消则返回 null</returns>
    Task<string?> SelectFolderAsync(string title);
}