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

### 第三方组件

以下组件非本项目开发，也不受本项目 GPL-3.0 协议约束，权利归各自所有者：

- [Shizuku](https://github.com/RikkaApps/Shizuku)：Android 权限桥接组件，随界面源码树附带了一份**未经修改**的官方 APK（`gui/MFAAvalonia/MFAAvalonia.Android/ThirdParty/Shizuku/shizuku.apk`），仅供界面的 Android 相关功能使用；本项目面向 Windows 桌面端，主流程不依赖它，不需要可自行删除。

## 💖 感谢贡献者

感谢所有为 MaaXA 以及MAAFW生态添砖加瓦的开发者们！🎉

[![贡献者](https://contrib.rocks/image?repo=AkumaYUC/MaaXA&max=1000)](https://github.com/AkumaYUC/MaaXA/graphs/contributors)

## 免责声明

- 本软件开源、免费，按现状（AS IS）提供，不作任何明示或默示的保证。使用本软件即视为已阅读并同意本节全部条款。
- **合规责任自负**：本软件通过界面识别与模拟键鼠的方式操作第三方软件。软件的用户协议、服务条款或所在单位的内部政策**可能禁止使用自动化脚本或模拟操作**，请在使用前自行确认。因使用本软件而产生的账号处置、服务中断、政策处罚或任何纠纷，均由使用者自行承担，与本软件及开发者无关。
- **单据务必人工复核**：本软件自动录入单据，可能因界面改版、识别误差、时序波动等原因出现错录、漏录。**每次任务完成后，请自行核对系统内的单据内容**再进行后续业务操作。因未复核、错录、漏录或重复录入造成的任何业务损失与数据错误，本软件及开发者概不负责。
- 在任何情况下，开发者不对任何直接或间接损失（包括但不限于业务损失、数据损失、利润损失）承担责任。

## 许可证

[GPL-3.0](LICENSE)
