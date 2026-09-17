<!-- markdownlint-disable MD033 MD041 -->
<p align="center">
  <img alt="LOGO" src="docs/images/logo.png" width="256" height="256" />
</p>

<div align="center">

# MaaXA

金蝶云自动化办公助手

</div>

基于 [MaaFramework](https://github.com/MaaXYZ/MaaFramework) 构建，界面使用 [MFAAvalonia](https://github.com/MaaXYZ/MFAAvalonia)。

## 功能

- **直接调拨单**：把 Excel 调拨单拖进界面，自动逐单录入金蝶云（物料编码、数量、调出/调入仓库与仓位、调出/调入库存组织、备注）。任务全部成功后自动归档单据副本，并按保留期清理过期文件。
- **启动应用**：自动拉起金蝶云客户端；若已有窗口或进程则跳过，不会重复启动。

## 下载使用

1. 到 [Releases](../../releases) 下载 `MaaXA-win-x86_64-<版本>.zip`
2. 解压到任意目录
3. 运行 `MFAAvalonia.exe`
4. 如果提示缺少运行库，运行目录下的 `DependencySetup_依赖库安装_win.bat`

**环境要求**：Windows 10 / 11 + 金蝶云星空客户端。

## 使用步骤

1. 启动金蝶云客户端（程序也会自动尝试拉起）
2. 把 Excel 调拨单拖进「单据文件」
3. 点「开始任务」

Excel 表头需与下列一致：物料编码、数量、调出仓库、调出仓位、调入仓库、调入仓位、调出库存组织、调入库存组织、备注。两个库存组织填在整单第一行即可、其余行留空；同一张单里填代码还是填名称请保持一致。列在表格里的先后顺序不限，列名必须一字不差。

## 从源码构建

需要 .NET 10 SDK。图形界面源码在 `gui/MFAAvalonia`，自动化资源与 agent 分别在 `assets/`、`agent/`。

## 鸣谢

本项目由 **[MaaFramework](https://github.com/MaaXYZ/MaaFramework)** 强力驱动，界面基于 **[MFAAvalonia](https://github.com/MaaXYZ/MFAAvalonia)**，项目骨架来自 **[MaaPracticeBoilerplate](https://github.com/MaaXYZ/MaaPracticeBoilerplate)** 模板。

## 💖 感谢贡献者

感谢所有为 MaaXA 以及MAAFW生态添砖加瓦的开发者们！🎉

[![贡献者](https://contrib.rocks/image?repo=AkumaYUC/MaaXA&max=1000)](https://github.com/AkumaYUC/MaaXA/graphs/contributors)

## 许可证

[GPL-3.0](LICENSE)
