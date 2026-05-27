# Changelog

All notable changes to **VMware vCenter Inspect** will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

### Planned

- PowerCLI 回退模式（补 REST 拿不到的 snapshot 大小 / Alarm 历史 / 单 host CPU/Mem 实时）
- 多 vCenter 批量模式（`-VCenter @('vc1','vc2','vc3')`）
- 基线对比（跟上次 findings 做 diff，只输出新增 / 已解决告警）
- Telegram / 飞书 / 钉钉告警推送（`-Notify` 参数）
- 配置文件（阈值从 `vcenter_inspect.config.json` 读取）

---

## [1.0.0] — 2026-05-25

### Added

#### 主脚本 `vcenter_inspect.ps1`（1250 行）

- 17 章节完整巡检报告：vCenter 概览 / Appliance 健康 / 网络 DNS / NTP 时间同步 / 访问入口 / TLS 证书 / VAMI 备份 / 拓扑 / ESXi 主机 / Datastore / Portgroup / VM 总览 / VM 列表 / VMware Tools 抽样 / Appliance 服务 / 总体建议 / 免责声明
- **Dual-mode REST API**：自动 fallback `/api/...`（v8）→ `/rest/com/vmware/cis/...`（v6），覆盖 vCenter 6.5 ~ 8.0
- **18 条 Findings 评估规则**，按 critical / warn / info 分级
- **动态总体建议**：按 findings 等级自动归类成短期 / 中期 / 长期 3 列
- **工程师风 HTML 模板**：扁平卡片 + 单色边框 + 蓝色 metadata banner + 章节编号 + 左侧 TOC scroll-spy（IntersectionObserver，无 jQuery）+ 打印样式（Ctrl+P 出 PDF）
- **8 个命令行参数**：`-VCenter` / `-Username` / `-Password` / `-Output` / `-ToolsSampleSize` / `-SkipToolsSample` / `-DebugDump` / `-Quiet`
- **自动重试机制**：对 0 / 5xx 瞬时错误指数退避（1s / 2s / 4s），扛 sts-idmd 抖动
- **错误诊断 hint**：401 / 403 / 5xx / 网络不可达 / 证书问题分别给出具体修复建议
- **session 自动清理**：脚本退出时主动 DELETE session，不留挂
- **UTF-8 中文支持**：手控 `[System.Net.HttpWebRequest]` + `[Text.Encoding]::UTF8`，绕过 PS 5.1 `Invoke-RestMethod` GBK 误解码

#### 配套转换器

- `html_to_md.ps1` — HTML → Markdown 转换器，针对本项目 HTML 结构正则解析，输出 ~15 KB
- `html_to_docx.ps1` — HTML → Word 转换器，用 Office Word COM 注入样式，输出 ~40 KB

#### 文档

- README.md — 完整使用文档（核心特性 / 快速开始 / 参数 / 章节说明 / Findings 规则 / 兼容性 / 自动化场景 / 开发踩坑 / 路线图）
- 脱敏 demo 报告样张 `report_demo_2026-05-25.html`
- README 截图：`docs/report-preview.png`（首屏 hero）+ `docs/report-preview-full.jpg`（整页超长效果）

### Tested

- vCenter 8.0.3 单 Cluster：2 ESXi / 33 VM / 6 Datastore，**4 秒**采集，42 KB 报告，捕获 11 条 findings
- vCenter 6.5 LTS 多 Cluster：25 ESXi / 39 Datastore / 38 TB，**~30 秒**采集，49 KB 报告，捕获 26 条 findings（6 critical / 15 warn / 5 info）
- Windows PowerShell 5.1（Win11 跳板机）+ PowerShell 7.4

### Known Limitations

REST API 不暴露的部分本工具不采集，需要 PowerCLI / vmodl SOAP 补充：

- 单 ESXi 主机 CPU / Memory / Build / Uptime / Maintenance Mode 实时数据（8.0 端点 deprecated）
- VM 快照大小（REST 可拿快照树，size 字段不暴露）
- Alarm / Event / 告警历史
- License 详情（`/api/vcenter/licensing/licenses` 在 8.0 返回 404）
- 性能历史曲线（stats API 仍是 preview）

7.0 / 7.0U3 路径已通过 dual-mode 自动匹配，缺真实环境实测验证。

---

[Unreleased]: https://github.com/Aidan-996/VMware_vCenter_Inspect/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/Aidan-996/VMware_vCenter_Inspect/releases/tag/v1.0.0
