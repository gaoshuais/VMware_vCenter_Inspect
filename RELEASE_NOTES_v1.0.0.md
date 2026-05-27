**PowerShell + REST API 写的 vCenter 一键巡检** — 零依赖，6.5 ~ 8.0 全兼容

---

### 核心能力

| 能力     | 说明                                                       |
| -------- | ---------------------------------------------------------- |
| 零依赖   | PowerShell 5.1+，不装 PowerCLI / pyvmomi                   |
| 17 章节  | 概览 / Health / 网络 / NTP / 证书 / 备份 / 拓扑 / VM ...   |
| 跨版本   | 6.5 / 6.7 / 7.0 / 8.0 dual-mode 自动适配                   |
| 动态告警 | 18 条 Findings 规则，归类短 / 中 / 长期建议                |
| 三种格式 | HTML（主）/ Word（Office COM）/ Markdown                   |
| 只读     | 全 GET，DELETE 仅用于注销 session，不改任何配置            |

### 实测性能

| 环境           | 规模                            | 耗时   | 报告  | Findings |
| -------------- | ------------------------------- | ------ | ----- | -------- |
| 8.0 单 Cluster | 2 ESXi / 33 VM                  | 4 s    | 42 KB | 11       |
| 6.5 多 Cluster | 25 ESXi / 39 Datastore / 38 TB  | ~30 s  | 49 KB | 26       |

### 快速上手

```powershell
git clone https://github.com/Aidan-996/VMware_vCenter_Inspect.git
cd VMware_vCenter_Inspect
.\vcenter_inspect.ps1 -VCenter 10.0.0.20 `
                      -Username administrator@vsphere.local `
                      -Password 'xxxx'
```

跑完打开 `report_*.html` 看效果。

### 后续路线

| 版本 | 主要内容                                          |
| ---- | ------------------------------------------------- |
| v1.1 | PowerCLI 回退 / 多 vCenter 批量 / 基线 diff       |
| v1.2 | Telegram 推送 / 配置文件化 / 历史趋势图           |

---

完整变更见 [CHANGELOG.md](https://github.com/Aidan-996/VMware_vCenter_Inspect/blob/main/CHANGELOG.md) · MIT License · 反馈 [提 Issue](https://github.com/Aidan-996/VMware_vCenter_Inspect/issues)
