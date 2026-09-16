using MFAAvalonia.Configuration;
using MFAAvalonia.Helper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace MFAAvalonia.Helper;

/// <summary>
/// 单据归档：任务全部成功后把拖入的 Excel 复制进归档目录（文件名前加执行日期），
/// 并按保留期清理旧副本。只删文件名匹配「8位日期_」格式的副本，目录里其他文件一律不碰。
/// </summary>
public static class OrderArchiveHelper
{
    //剥旧前缀：兼容本 GUI 的 yyyyMMdd_ 和 agent 旧版归档的 yyyyMMdd-HHmmss_
    private static readonly Regex DatePrefixRegex = new(@"^(?:\d{8}|\d{8}-\d{6})_", RegexOptions.Compiled);

    //只删「8位日期_原文件名」格式的副本，其他文件（用户手动放的东西）不碰
    private static readonly Regex ArchivedCopyRegex = new(@"^\d{8}_.+\.(xlsx|xlsm)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>
    /// 归档一份处理完成的单据副本；返回归档后的完整路径，失败返回 null（不抛异常，失败只记日志不拦收尾流程）。
    /// 同一文件重复执行：先剥旧日期前缀再加今日日期 → 文件名恒定，File.Copy 覆盖旧副本。
    /// </summary>
    public static string? ArchiveOrderFile(string sourcePath, string archiveDirectory)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
            {
                LoggerHelper.Warning($"[单据归档] 源文件不存在，跳过归档: {sourcePath}");
                return null;
            }
            if (string.IsNullOrWhiteSpace(archiveDirectory))
            {
                LoggerHelper.Warning("[单据归档] 归档目录为空，跳过归档");
                return null;
            }

            Directory.CreateDirectory(archiveDirectory);

            var fileName = Path.GetFileName(sourcePath);
            var bareName = DatePrefixRegex.Replace(fileName, string.Empty);  //剥掉旧日期前缀（若有）
            var datedName = $"{DateTime.Now:yyyyMMdd}_{bareName}";
            var destPath = Path.Combine(archiveDirectory, datedName);

            File.Copy(sourcePath, destPath, overwrite: true);
            LoggerHelper.Info($"[单据归档] 已归档到 {destPath}");
            return destPath;
        }
        catch (Exception e)
        {
            LoggerHelper.Error($"[单据归档] 归档失败: {e.Message}", e);
            return null;
        }
    }

    /// <summary>
    /// 按保留期清理归档目录里的过期副本。retentionDays &lt;= 0 时不清理。
    /// 日期取自文件名前缀（不是文件修改时间）——文件被复制后 mtime 是复制时刻，语义会漂。
    /// </summary>
    public static int CleanupExpired(string archiveDirectory, int retentionDays)
    {
        if (retentionDays <= 0 || string.IsNullOrWhiteSpace(archiveDirectory) || !Directory.Exists(archiveDirectory))
            return 0;

        var deleted = 0;
        try
        {
            var deadline = DateTime.Today.AddDays(-retentionDays);
            foreach (var file in Directory.EnumerateFiles(archiveDirectory))
            {
                try
                {
                    var name = Path.GetFileName(file);
                    if (!ArchivedCopyRegex.IsMatch(name)) continue;  //不匹配副本格式的文件一律不碰

                    var datePart = name[..8];  //前 8 位就是 yyyyMMdd
                    if (!DateTime.TryParseExact(datePart, "yyyyMMdd",
                            System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.None, out var runDate))
                        continue;

                    if (runDate < deadline)
                    {
                        File.Delete(Path.Combine(archiveDirectory, name));
                        deleted++;
                        LoggerHelper.Info($"[单据归档] 已清理过期副本: {name}");
                    }
                }
                catch (Exception e)
                {
                    LoggerHelper.Error($"[单据归档] 清理单个文件失败: {file}: {e.Message}");
                }
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Error($"[单据归档] 清理归档目录失败: {e.Message}", e);
        }
        return deleted;
    }

    /// <summary>
    /// 读实例配置里的保留期设置，返回保留天数；0 = 不清理。
    /// retention 取值：off / week / month / half_year / custom
    /// </summary>
    public static int ResolveRetentionDays(InstanceConfiguration instanceConfig)
    {
        var retention = instanceConfig.GetValue(ConfigurationKeys.OrderArchiveRetention, "off");
        if (retention == "custom")
        {
            var days = instanceConfig.GetValue(ConfigurationKeys.OrderArchiveCustomDays, 0);
            return Math.Max(0, days);
        }
        return retention switch
        {
            "week" => 7,
            "month" => 30,
            "half_year" => 180,
            _ => 0,
        };
    }

    /// <summary>
    /// 归档目录：用户在 option 里填了就用用户的；没填就用发行根下的「已处理单据」。
    /// </summary>
    public static string ResolveArchiveDirectory(InstanceConfiguration instanceConfig)
    {
        var configured = instanceConfig.GetValue(ConfigurationKeys.OrderArchiveDirectory, string.Empty);
        if (!string.IsNullOrWhiteSpace(configured)) return configured;
        return Path.Combine(AppPaths.InstallRoot, "已处理单据");
    }
}
