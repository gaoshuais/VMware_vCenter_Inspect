## VMware vCenter Inspect v1.0.0 — 首发

PowerShell + REST API 写的 vCenter 一键巡检工具。一条命令 4 秒出一份工程师风 HTML 报告，零依赖，vCenter 6.5 / 6.7 / 7.0 / 8.0 全兼容。

### 核心能力

- **零依赖** — 只需 Windows PowerShell 5.1+，不装 PowerCLI / pyvmomi，70 KB 单文件
- **Dual-mode REST API** — 自动 fallback `/api/...`（v8）→ `/rest/com/vmware/cis/...`（v6），覆盖 6.5 ~ 8.0
- **17 章节报告** — vCenter 概览 / Health / 网络 / NTP / 访问 / 证书 / 备份 / 拓扑 / 主机 / Datastore / 网络 / VM / Tools / 服务 / 动态建议 / 免责
- **18 条 Findings 评估规则** — critical / warn / info 三级，自动归类成短中长期三栏建议
- **三种输出格式** — HTML（主输出，扁平卡片+scroll-spy 侧栏+打印样式）+ DOCX（Word COM）+ Markdown
- **只读采集** — 全部 GET，DELETE 仅用于注销 session，不修改 vCenter 任何配置

### 实测性能

| 环境 | 规模 | 耗时 | 报告体积 | Findings |
|---|---|---|---|---|
| 8.0.3 单 Cluster | 2 ESXi / 33 VM | 4 秒 | 42 KB | 11 |
| 6.5 LTS 多 Cluster | 25 ESXi / 39 Datastore / 38 TB | ~30 秒 | 49 KB | 26 (6c/15w/5i) |

### 一句话上手

```powershell
git clone https://github.com/Aidan-996/VMware_vCenter_Inspect.git
cd VMware_vCenter_Inspect
.\vcenter_inspect.ps1 -VCenter 10.0.0.20 -Username administrator@vsphere.local -Password 'xxxx'
```

跑完打开 `report_10.0.0.20_<date>.html`，就这样。

### 看效果

- 首屏：`docs/report-preview.png`
- 整页：`docs/report-preview-full.jpg`
- 可执行报告：`report_demo_2026-05-25.html`（浏览器打开看交互效果）

### Known Issues

- 7.0 / 7.0U3 路径已通过 dual-mode 自动匹配，缺真实环境实测，欢迎社区反馈
- REST API 不暴露：VM 快照大小 / Alarm 历史 / 单 host 实时负载 / License / 性能曲线（详见 README）
- Linux + pwsh 7 理论可用，未实测

### 路线图（v1.1 计划中）

- PowerCLI 回退模式（补 REST 拿不到的数据）
- 多 vCenter 批量
- 基线对比（findings diff）
- Telegram / 飞书 / 钉钉推送

完整变更见 [CHANGELOG.md](https://github.com/Aidan-996/VMware_vCenter_Inspect/blob/main/CHANGELOG.md)。

---

**MIT License** · 问题 / 建议 / 兼容性反馈，[提 Issue](https://github.com/Aidan-996/VMware_vCenter_Inspect/issues)
