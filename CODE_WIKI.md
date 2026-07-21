# VMware vCenter Inspect - Code Wiki

> **版本**: v1.2.0  
> **语言**: PowerShell 5.1+ / PowerShell 7+  
> **许可证**: MIT  
> **核心功能**: vCenter 一键巡检工具，通过 REST API + 可选 PowerCLI 回退生成工程师风 HTML 报告

---

## 目录

1. [项目概述](#1-项目概述)
2. [整体架构](#2-整体架构)
3. [文件结构](#3-文件结构)
4. [核心模块详解](#4-核心模块详解)
   - 4.1 [全局初始化模块](#41-全局初始化模块)
   - 4.2 [日志输出模块](#42-日志输出模块)
   - 4.3 [HTTP 调用层](#43-http-调用层)
   - 4.4 [PowerCLI 回退层](#44-powercli-回退层)
   - 4.5 [数据采集层](#45-数据采集层)
   - 4.6 [评估告警层](#46-评估告警层)
   - 4.7 [HTML 渲染层](#47-html-渲染层)
5. [关键函数说明](#5-关键函数说明)
6. [数据流与依赖关系](#6-数据流与依赖关系)
7. [命令行参数参考](#7-命令行参数参考)
8. [运行方式](#8-运行方式)
9. [开发约定与踩坑记录](#9-开发约定与踩坑记录)
10. [版本演进路线](#10-版本演进路线)

---

## 1. 项目概述

### 1.1 项目定位

VMware vCenter Inspect 是一款零依赖的 vCenter 一键巡检工具。使用 PowerShell + vSphere REST API 实现，一条命令约 4 秒即可生成一份包含 19 个章节的工程师风 HTML 巡检报告。

### 1.2 核心特性

| 特性 | 说明 |
|---|---|
| **零依赖** | 只需 Windows PowerShell 5.1+，无需安装 PowerCLI / Python / 任何第三方模块 |
| **Dual-mode REST** | 自动适配 vCenter 6.5/6.7（`/rest/com/vmware/cis/...`）与 7.0/8.0+（`/api/...`）两套 API |
| **19 章节** | 概览 / Health / 网络 / NTP / 访问 / 证书 / 备份 / 拓扑 / 主机 / Datastore / 网络 / VM 总览 / VM 列表 / Tools / 服务 / 快照 / 告警 / 总体建议 / 免责 |
| **动态告警** | 28 条评估规则按 critical / warn / info 分级，自动归类为短期 / 中期 / 长期建议 |
| **三种格式** | HTML（主输出）+ DOCX（Word COM 转换）+ Markdown（正则转换） |
| **4 套主题** | light / dark / minimal / amber，报告内可实时切换，localStorage 持久化 |
| **只读采集** | REST API 全为 GET，DELETE 仅用于注销 session，不修改任何配置 |

### 1.3 兼容性矩阵

| 组件 | 支持版本 |
|---|---|
| **vCenter** | 6.5 / 6.7 / 7.0 / 8.0 / 8.0U2 / 8.0U3 |
| **PowerShell** | Windows PowerShell 5.1+ / PowerShell 7.0+ |
| **操作系统** | Windows 10/11 / Windows Server 2016+ |
| **PowerCLI（可选）** | VMware.PowerCLI 13.0+ |

---

## 2. 整体架构

### 2.1 架构分层图

```
┌─────────────────────────────────────────────────────────┐
│                    命令行入口层                           │
│  param() 参数解析 → 主流程 try/catch → exit code         │
├─────────────────────────────────────────────────────────┤
│                    日志输出层                             │
│  Log-Banner / Log-Step / Log-Info / Log-Warn / Log-Err   │
├─────────────────────────────────────────────────────────┤
│                    HTTP 调用层 (Dual-mode)                │
│  Invoke-VCRaw → Invoke-VCRetry → Resolve-VCPath → Invoke-VC │
│  v8 (/api/...)  ←自动探测→  v6 (/rest/...)                │
├─────────────────────────────────────────────────────────┤
│                    数据采集层                             │
│  Collect-All (REST 14 步)  +  Collect-PowerCLI (3 步)    │
├─────────────────────────────────────────────────────────┤
│                    评估告警层                             │
│  Eval-Findings → 28 条规则 → critical/warn/info 分级    │
├─────────────────────────────────────────────────────────┤
│                    HTML 渲染层                            │
│  Render-Report → 19 章节 + 4 主题 + scroll-spy + 切换器  │
├─────────────────────────────────────────────────────────┤
│                    格式转换层 (独立脚本)                   │
│  html_to_docx.ps1 (Word COM)  |  html_to_md.ps1 (正则)   │
└─────────────────────────────────────────────────────────┘
```

### 2.2 核心设计原则

1. **零依赖原则**: 不强制安装任何第三方模块，PowerCLI 仅作为可选增强
2. **只读原则**: 全部 GET 请求，DELETE 仅用于 session 注销
3. **优雅降级**: 不支持的 API 端点自动 skip，对应章节显示 N/A，不打断整体流程
4. **自动重试**: 对网络错误和 5xx 瞬时错误采用指数退避（1s/2s/4s）
5. **错误诊断**: 不同错误码（401/403/404/5xx/0）给出具体的修复建议

---

## 3. 文件结构

```
/workspace/
├── vcenter_inspect.ps1          # 主巡检脚本 (≈1900 行)
├── html_to_md.ps1               # HTML → Markdown 转换器 (≈320 行)
├── html_to_docx.ps1             # HTML → Word 转换器 (≈90 行)
├── README.md                    # 使用文档
├── CHANGELOG.md                 # 版本变更日志
├── RELEASE_NOTES_v1.0.0.md      # v1.0.0 发布说明
├── LICENSE                      # MIT 许可证
└── .gitignore                   # Git 忽略规则
```

### 3.1 主脚本文件职责

| 文件 | 职责 | 行数 | 依赖 |
|---|---|---|---|
| [vcenter_inspect.ps1](file:///workspace/vcenter_inspect.ps1) | 核心巡检逻辑，数据采集 + 评估 + HTML 渲染 | ~1900 | 无（PowerCLI 可选） |
| [html_to_docx.ps1](file:///workspace/html_to_docx.ps1) | HTML → DOCX 格式转换，依赖 Word COM | ~90 | Microsoft Word (COM) |
| [html_to_md.ps1](file:///workspace/html_to_md.ps1) | HTML → Markdown 格式转换，纯正则解析 | ~320 | 无 |

---

## 4. 核心模块详解

### 4.1 全局初始化模块

**位置**: [vcenter_inspect.ps1#L76-L99](file:///workspace/vcenter_inspect.ps1#L76-L99)

#### 功能概述
脚本启动时的全局环境配置，确保在不同 PowerShell 版本和系统环境下稳定运行。

#### 关键配置项

| 配置项 | 值 | 说明 |
|---|---|---|
| `$ErrorActionPreference` | `'Stop'` | 错误立即终止，避免静默失败 |
| `$ProgressPreference` | `'SilentlyContinue'` | 禁用进度条，提升性能 |
| `[Console]::OutputEncoding` | UTF-8 | 终端输出编码，避免中文乱码 |
| `$OutputEncoding` | UTF-8 | 管道输出编码 |
| `SecurityProtocol` | Tls12 + Tls11 + Tls | 兼容不同 vCenter 的 TLS 版本 |
| `TrustAllCertsPolicy` | 自定义类 | 绕过 vCenter 默认自签证书校验 |

#### TrustAllCertsPolicy 类
通过 `Add-Type` 动态注入 C# 类，实现 `ICertificatePolicy` 接口，使 `HttpWebRequest` 信任所有 SSL 证书。这是连接自签 vCenter 的标准做法。

---

### 4.2 日志输出模块

**位置**: [vcenter_inspect.ps1#L101-L127](file:///workspace/vcenter_inspect.ps1#L101-L127)

#### 全局变量

| 变量 | 类型 | 初始值 | 说明 |
|---|---|---|---|
| `$Script:StepIdx` | int | 0 | 当前步骤序号 |
| `$Script:StepTotal` | int | 14 | REST 采集总步数 |
| `$Script:T0` | DateTime | Get-Date | 脚本启动时间 |
| `$Script:PCLIEnabled` | bool | $false | PowerCLI 回退是否启用 |

#### 函数列表

| 函数名 | 参数 | 功能 |
|---|---|---|
| `Log-Banner` | 无 | 打印启动横幅（版本、目标、时间、步骤数、用户名） |
| `Log-Step` | `[string]$msg` | 打印步骤进度（含序号、百分比） |
| `Log-Info` | `[string]$msg` | 打印灰色信息文本 |
| `Log-Warn` | `[string]$msg` | 打印黄色警告文本 |
| `Log-Err` | `[string]$msg` | 打印红色错误文本 |

#### 输出格式示例
```
═══════════════════════════════════════════════════════════
 vCenter Inspect  v1.2.0 |  target: 10.0.0.20   |  2026-05-25 11:49:32
 steps: 14                  |  user  : administrator@vsphere.local
═══════════════════════════════════════════════════════════

  [ 1/14] (  7%) Login vCenter ...
  [ 2/14] ( 14%) 1. 系统版本与运行时间 ...
```

---

### 4.3 HTTP 调用层

**位置**: [vcenter_inspect.ps1#L128-L317](file:///workspace/vcenter_inspect.ps1#L128-L317)

#### 核心设计: Dual-mode API 适配

HTTP 层是整个项目最核心的技术模块，实现了 vCenter 6.5 ~ 8.0 的自动兼容。

#### 全局状态变量

| 变量 | 类型 | 说明 |
|---|---|---|
| `$Script:VCBase` | string | vCenter 基础 URL `https://<vc>` |
| `$Script:Session` | string | API session token |
| `$Script:ApiStyle` | string | `'v8'` (7.0+) 或 `'v6'` (6.5/6.7) |
| `$Script:DebugLog` | string | Debug 日志文件路径 |
| `$Script:V6Unsupported` | array | v6 模式不支持的端点列表 |

#### 函数调用链

```
Invoke-VC (高层封装, 返回 PSObject)
    ↓
Resolve-VCPath (路径转换: v8 → v6)
    ↓
Invoke-VCRetry (带指数退避重试)
    ↓
Invoke-VCRaw (底层 HttpWebRequest 封装, 返回 {Code, Text, Error})
```

#### 4.3.1 Invoke-VCRaw — 底层 HTTP 封装

**位置**: [vcenter_inspect.ps1#L168-L212](file:///workspace/vcenter_inspect.ps1#L168-L212)

**核心技术点**:
- 使用 `[System.Net.HttpWebRequest]` 而非 `Invoke-RestMethod`，手动控制 UTF-8 编码，避免 PS 5.1 的 GBK 乱码问题
- 支持 `vmware-api-session-id` header 认证
- 统一错误捕获，返回结构化结果 `{Code, Text, Error}`

**参数**:
| 参数 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `Path` | string | - | API 路径，如 `/api/appliance/system/version` |
| `Method` | string | `'GET'` | HTTP 方法 |
| `AuthBasic` | string | - | Base64 编码的 Basic Auth（登录用） |
| `Body` | string | - | 请求体 JSON |

**返回值**: `[pscustomobject]@{ Code = [int]; Text = [string]; Error = [string] }`

#### 4.3.2 Invoke-VCRetry — 指数退避重试

**位置**: [vcenter_inspect.ps1#L214-L232](file:///workspace/vcenter_inspect.ps1#L214-L232)

**重试策略**:
- **重试条件**: HTTP code = 0（网络错误）或 5xx（服务端瞬时错误）
- **不重试**: 4xx（认证失败 / 资源不存在 / 客户端错误）
- **重试间隔**: 指数退避 `2^(attempt-1)` 秒 → 1s / 2s / 4s
- **最大重试**: 默认 3 次

#### 4.3.3 Resolve-VCPath — v8 → v6 路径转换

**位置**: [vcenter_inspect.ps1#L234-L247](file:///workspace/vcenter_inspect.ps1#L234-L247)

**转换规则**:
1. v8 模式（`ApiStyle -ne 'v6'`）: 直接返回原路径
2. v6 模式:
   - `/api/session` → `/rest/com/vmware/cis/session`
   - VM 子资源（`/api/vcenter/vm/*/tools` 等）→ 返回 `$null`（直接跳过）
   - 匹配 `$V6Unsupported` 列表 → 返回 `$null`（直接跳过）
   - 其他: `/api/...` → `/rest/...`

#### 4.3.4 Invoke-VC — 高层业务封装

**位置**: [vcenter_inspect.ps1#L249-L265](file:///workspace/vcenter_inspect.ps1#L249-L265)

**功能**:
- 调用 `Resolve-VCPath` 转换路径
- 不支持的端点直接返回 `$null`
- 成功时 `ConvertFrom-Json` 解析
- v6 模式自动 unwrap `{"value": ...}` 包装
- 失败时返回 `$null`（不抛异常，优雅降级）

#### 4.3.5 Login-VC — 双模式登录探测

**位置**: [vcenter_inspect.ps1#L267-L310](file:///workspace/vcenter_inspect.ps1#L267-L310)

**登录流程**:
1. 构造 Basic Auth header（用户名:密码 Base64）
2. **先试 v8**: `POST /api/session`，2 次重试
3. v8 成功 → 设置 `ApiStyle = 'v8'`，返回 session id
4. v8 失败 → fallback v6: `POST /rest/com/vmware/cis/session`，3 次重试
5. v6 成功 → 设置 `ApiStyle = 'v6'`，解析 `{value: "<sid>"}`
6. 都失败 → 精确诊断错误并抛出

**错误诊断映射**:
| HTTP Code | 诊断提示 |
|---|---|
| 401 | 用户名/密码错误，或账号已锁定（SSO 默认 5 次失败锁 5 分钟） |
| 403 | 账号无 API 权限 |
| 500+ | vCenter 服务端错误（sts-idmd / vapi-endpoint 可能异常），建议重启 vmware-vapi-endpoint |
| 0 | 网络不可达，检查路由 / 防火墙 / TLS |
| 404 | 该 vCenter REST API 端点都不可用，可能 < 6.5 版本 |

#### 4.3.6 Logout-VC — 会话清理

**位置**: [vcenter_inspect.ps1#L311-L316](file:///workspace/vcenter_inspect.ps1#L311-L316)

**功能**: `DELETE` 当前 session，不留挂 vCenter session 资源。

---

### 4.4 PowerCLI 回退层

**位置**: [vcenter_inspect.ps1#L356-L535](file:///workspace/vcenter_inspect.ps1#L356-L535)

#### 设计目的
REST API 8.0 仍未暴露的三类关键数据，通过可选 PowerCLI 回退补全：
1. 单 ESXi 主机实时数据（CPU/Mem/Uptime/Build）
2. VM 快照清单（含大小 + 年龄 + 链深度）
3. 当前 Triggered Alarms

#### 启用策略
- **自动检测**: 检测到 `VMware.VimAutomation.Core` 模块即自动启用
- **强制启用**: `-UsePowerCLI` 参数
- **强制跳过**: `-SkipPowerCLI` 参数

#### 函数列表

| 函数名 | 功能 |
|---|---|
| `Try-LoadPowerCLI` | 检测并加载 PowerCLI 模块，配置证书忽略 / CEIP 关闭 / Single 模式 |
| `Connect-PowerCLI` | 使用 `Connect-VIServer` 连接 vCenter |
| `Disconnect-PowerCLI` | 断开 PowerCLI 连接 |
| `Collect-PowerCLI` | 主入口，依次采集 Host 实时 / 快照 / 告警三类数据 |

#### 4.4.1 Collect-PowerCLI 三阶段采集

**阶段 1: 单 Host 实时数据** (Get-VMHost)
- 字段: Name, ConnState, PowerState, NumCpu, CpuUsedMhz, CpuTotalMhz, CpuPct, MemUsedGB, MemTotalGB, MemPct, UptimeDays, Version, Build, Vendor, Model
- 存入 `$Data.PCLI_Hosts`

**阶段 2: VM 快照清单** (Get-Snapshot)
- 全 VM 遍历，收集每个快照的: VM, Snapshot, Created, AgeDays, SizeGB, PowerState, Desc
- 聚合统计: VmCount, Count, TotalGB, OldestDays, MaxChain, Top10
- 存入 `$Data.PCLI_Snapshots`

**阶段 3: 当前 Triggered Alarms** (TriggeredAlarmState)
- 遍历实体类型: Datacenter / Cluster / VMHost / Datastore / VM
- 字段: Entity, EntityType, Alarm, Status, Time, AgeHours, Acked
- 存入 `$Data.PCLI_Alarms`

---

### 4.5 数据采集层

**位置**: [vcenter_inspect.ps1#L538-L658](file:///workspace/vcenter_inspect.ps1#L538-L658)

#### 数据容器: $Data
使用 `[ordered]@{}` 有序哈希表存储所有采集数据，键名即章节名。

#### Collect-All — 14 步 REST 采集流程

| 步骤 | 章节 | API 端点 | 存入 $Data |
|---|---|---|---|
| 0 | 登录 | `POST /api/session` | `$Script:Session` |
| 1 | 系统版本与运行时间 | `/api/appliance/system/version` | `.SysVersion` |
| | | `/api/appliance/system/uptime` | `.SysUptime` |
| | | `/api/appliance/system/time` | `.SysTime` |
| | | `/api/vcenter/deployment` | `.Deployment` |
| | | `/api/appliance/update` | `.Update` |
| 2 | Appliance 健康 | `/api/appliance/health/{system,storage,mem,swap,load,database-storage,applmgmt,software-packages,system/lastcheck}` | `.Health` |
| 3 | 网络 / DNS / NTP / 时间同步 | `/api/appliance/networking` | `.Networking` |
| | | `/api/appliance/networking/interfaces` | `.NetInterfaces` |
| | | `/api/appliance/ntp` | `.NTP` |
| | | `/api/appliance/timesync` | `.Timesync` |
| 4 | 访问入口 | `/api/appliance/access/{ssh,shell,dcui,consolecli}` | `.Access` |
| 5 | vCenter 证书 | `/api/vcenter/certificate-management/vcenter/tls` | `.CertTLS` |
| 6 | 备份策略与历史 | `/api/appliance/recovery/backup/job` | `.BackupJobs` |
| | | `/api/appliance/recovery/backup/schedules` | `.BackupSchedules` |
| 7 | 拓扑 | `/api/vcenter/datacenter` | `.Datacenters` |
| | | `/api/vcenter/cluster` | `.Clusters` |
| | | `/api/vcenter/folder` | `.Folders` |
| | | `/api/vcenter/resource-pool` | `.ResPools` |
| 8 | ESXi 主机 | `/api/vcenter/host` | `.Hosts` |
| 9 | Datastore | `/api/vcenter/datastore` | `.Datastores` |
| 10 | 网络 | `/api/vcenter/network` | `.Networks` |
| 11 | 虚拟机清单 | `/api/vcenter/vm` | `.VMs` |
| 12 | Appliance 服务列表 | `/api/appliance/services` | `.Services` |
| 13 | VMware Tools 抽样 | `/api/vcenter/vm/{vm}/tools` | `.ToolsSample` |
| 14 | 关闭 session | `DELETE /api/session` | - |

#### VMware Tools 抽样策略
- 仅对 `power_state = POWERED_ON` 的 VM 抽样
- 默认抽样数量: 16（`-ToolsSampleSize` 可调）
- 抽样字段: VM, VMId, RunState, Version, Status, Install, Policy
- 可通过 `-SkipToolsSample` 跳过以节省时间

---

### 4.6 评估告警层

**位置**: [vcenter_inspect.ps1#L660-L793](file:///workspace/vcenter_inspect.ps1#L660-L793)

#### Eval-Findings — 动态告警评估

**输出结构**:
```powershell
[pscustomobject]@{
    Severity = 'critical' | 'warn' | 'info'
    Area     = '健康' | '时间同步' | 'DNS' | '访问' | '证书' | '备份' | ...
    Title    = '问题标题'
    Detail   = '详细说明与建议'
}
```

#### 评估规则汇总（共 28 条）

##### REST 维度（18 条）

| 维度 | 规则 | 级别 |
|---|---|---|
| Appliance Health | 任意项 = yellow/orange | warn |
| Appliance Health | 任意项 = red | critical |
| NTP | 服务器列表为空 | warn |
| Timesync | 状态 ≠ NTP | warn |
| DNS | hostname = localhost 或空 | warn |
| DNS | server 数 < 2 | info |
| SSH | 已启用 | warn |
| DCUI | 已启用 | info |
| TLS 证书 | 剩余 < 30 天 | critical |
| TLS 证书 | 剩余 < 90 天 | warn |
| TLS 证书 | 自签（issuer 含 localhost/VMSCA） | info |
| VAMI 备份 | jobs 为空 | warn |
| VAMI 备份 | schedules 为空 | warn |
| Cluster | ≥ 2 节点 + HA 关 | warn |
| Cluster | ≥ 2 节点 + DRS 关 | info |
| Datastore | 使用率 ≥ 90% | critical |
| Datastore | 使用率 ≥ 80% | warn |
| Datastore | 使用率 ≥ 70% | info |
| VMware Tools | 抽样有 OLD 版本 | info |
| vCenter update | 非 UP_TO_DATE | info |

##### PowerCLI 维度（10 条）

| 维度 | 规则 | 级别 |
|---|---|---|
| ESXi Memory | 使用率 ≥ 90% | critical |
| ESXi Memory | 使用率 ≥ 80% | warn |
| ESXi CPU | 使用率 ≥ 85% | warn |
| ESXi 连接 | Disconnected / NotResponding | critical |
| ESXi Uptime | > 365 天 | info |
| Snapshot | 最老快照 > 90 天 | critical |
| Snapshot | 最老快照 > 30 天 | warn |
| Snapshot | 单 VM 链深 > 3 层 | warn |
| Snapshot | 总占用 > 1 TB | info |
| Alarm | RED 告警 > 0 | critical |
| Alarm | YELLOW 告警 > 5 | warn |

#### 建议自动归类
在 HTML 渲染阶段，按 Severity 自动归类到三列：
- **短期（立即处理）**: critical
- **中期（1-2 周内）**: warn
- **长期（持续改进）**: info

---

### 4.7 HTML 渲染层

**位置**: [vcenter_inspect.ps1#L795-L1856](file:///workspace/vcenter_inspect.ps1#L795-L1856)

#### Render-Report — 报告生成入口

**渲染流程**:
1. 预处理: 从 `$Data` 提取常用变量，统计 VM/Datastore/Service 汇总数据
2. 评估: 调用 `Eval-Findings` 生成告警列表
3. 构造 CSS: 4 套主题 palette + 业务样式 + 打印样式
4. 构造 HTML: banner → summary cards → 19 个 section → footer → JS
5. 返回完整 HTML 字符串

#### 4.7.1 主题系统（v1.2+）

**CSS 变量架构**: 使用 `:root` / `[data-theme="..."]` 选择器，~15 个 CSS 变量集中控制全套配色。

| 变量 | 用途 |
|---|---|
| `--bg` / `--bg-side` / `--bg-card` / `--bg-card-2` | 四级背景色 |
| `--fg` / `--fg-mute` / `--fg-dim` | 三级前景文字色 |
| `--border` / `--border-strong` | 边框色 |
| `--accent` / `--accent-2` | 主色调 / 副色调 |
| `--green` / `--amber` / `--red` / `--gray` / `--blue` | 状态色 |
| `--*-bg` | 状态色的半透明背景（用于 badge / icon） |

**4 套内置主题**:

| 主题 | 风格 | Accent 色 |
|---|---|---|
| `light`（默认） | 白底 + 工程师蓝 | `#1565c0` |
| `dark` | 深灰 + 亮蓝 NOC | `#58a6ff` |
| `minimal` | 灰白 + 近黑 accent | `#27272a` |
| `amber` | 米色 + 琥珀棕 | `#b45309` |

**主题切换机制**:
- 生成时: `<html data-theme="$Theme">` 设置初始值
- 运行时: 右上角 4 色圆点切换器，点击即时切换
- 持久化: `localStorage['vci_theme']` 保存用户偏好
- 优先级: localStorage > `-Theme` 参数 > 默认 `light`

#### 4.7.2 页面结构

```
<body>
├── <div class="theme-switch">    # 主题切换器（右上角悬浮）
├── <aside class="sidebar">       # 左侧导航栏
│   ├── brand (logo + 版本号)
│   └── nav (19 个章节链接)
└── <main class="main">
    ├── <div class="banner">      # 顶部 metadata 横幅
    ├── <div class="cards">       # 6 个 Summary 卡片
    ├── <section id="sec-overview">   # 1. vCenter 概览
    ├── <section id="sec-health">     # 2. Appliance 健康
    ├── ... (共 19 个 section)
    └── <footer>                  # 页脚（版本 / 耗时 / 生成时间）
</main>
<script>  # scroll-spy + 主题切换 JS
```

#### 4.7.3 19 个章节结构

| # | ID | 标题 | 数据源 | 组件类型 |
|---|---|---|---|---|
| 1 | `sec-overview` | vCenter 概览 | `.SysVersion` + `.Update` | info-grid |
| 2 | `sec-health` | Appliance 健康 | `.Health` | info-grid |
| 3 | `sec-network` | 网络与 DNS | `.Networking` + `.NetInterfaces` | info-grid + table |
| 4 | `sec-ntp` | NTP 与时间同步 | `.NTP` + `.Timesync` | info-grid |
| 5 | `sec-access` | 访问入口 | `.Access` | info-grid |
| 6 | `sec-cert` | vCenter TLS 证书 | `.CertTLS` | info-grid |
| 7 | `sec-backup` | VAMI 备份策略 | `.BackupJobs` + `.BackupSchedules` | info-grid |
| 8 | `sec-topo` | Datacenter/Cluster/Folder | `.Datacenters` + `.Clusters` + `.Folders` + `.ResPools` | table |
| 9 | `sec-hosts` | ESXi 主机 | `.Hosts` + `.PCLI_Hosts` | table (+ 实时副表) |
| 10 | `sec-ds` | Datastore | `.Datastores` | info-grid + table (带进度条) |
| 11 | `sec-net` | 网络 / Portgroup | `.Networks` | info-grid + table |
| 12 | `sec-vm-sum` | VM 总览 | `.VMs` 聚合统计 | info-grid |
| 13 | `sec-vm-list` | VM 列表 | `.VMs` 排序 | table |
| 14 | `sec-tools` | VMware Tools 抽样 | `.ToolsSample` | table |
| 15 | `sec-services` | Appliance 服务 | `.Services` (核心 20 项) | info-grid + table |
| 16 | `sec-snap` | VM 快照健康 | `.PCLI_Snapshots` | info-grid + Top10 table |
| 17 | `sec-alarm` | Alarm 当前告警 | `.PCLI_Alarms` | info-grid + table |
| 18 | `sec-find` | 总体建议 | `Eval-Findings` 结果 | 3 列 findings |
| 19 | `sec-discl` | 免责声明 | 静态文本 | disclaimer 块 |

#### 4.7.4 核心 UI 组件

| 组件类名 | 用途 |
|---|---|
| `.banner` | 顶部蓝色 metadata 横幅 |
| `.cards` / `.card` | 6 个概览卡片（网格布局） |
| `.info-grid` | 键值对信息网格（dt/dd 两列） |
| `table` | 数据表格（斑马纹 + hover 高亮） |
| `.badge` | 状态标签（green/amber/red/gray/blue） |
| `.bar` | 使用率进度条（green/amber/red 三级） |
| `.findings` / `.find-col` | 总体建议三列布局（short/mid/long） |
| `.disclaimer` | 免责声明灰色块 |
| `.sidebar` | 左侧固定导航栏（带 scroll-spy） |
| `.theme-switch` | 右上角主题切换器 |

#### 4.7.5 JavaScript 功能

**Scroll Spy（滚动高亮）**:
- 使用 `IntersectionObserver` 监听 section 可见性
- 滚动到哪个章节，左侧导航对应项高亮
- 观察器 margin: `-40% 0px -55% 0px`（中间 5% 区域触发）

**主题切换**:
- 启动时读取 `localStorage['vci_theme']` 覆盖初始主题
- 点击切换器按钮时设置 `<html data-theme>` + 写 localStorage
- 按钮 `.active` 类同步当前选中状态

**打印样式** (`@media print`):
- 隐藏侧边栏和主题切换器
- 主体占满宽度，背景变白，边框变灰色

---

## 5. 关键函数说明

### 5.1 辅助工具函数

| 函数名 | 位置 | 功能 |
|---|---|---|
| `Format-Bytes` | [L321-L326](file:///workspace/vcenter_inspect.ps1#L321-L326) | 字节数格式化为可读单位（B/KB/MB/GB/TB/PB） |
| `Format-DaysFromSec` | [L327-L332](file:///workspace/vcenter_inspect.ps1#L327-L332) | 秒数格式化为 "X 天 Y 小时" |
| `Html-Encode` | [L333-L337](file:///workspace/vcenter_inspect.ps1#L333-L337) | HTML 转义，防止 XSS |
| `Badge` | [L338-L341](file:///workspace/vcenter_inspect.ps1#L338-L341) | 生成 badge HTML 片段 |
| `Health-Badge` | [L342-L353](file:///workspace/vcenter_inspect.ps1#L342-L353) | 健康状态映射到 badge 颜色（green/yellow/orange/red/gray） |
| `Write-Dump` | [L145-L151](file:///workspace/vcenter_inspect.ps1#L145-L151) | Debug 模式下写入请求日志 |

### 5.2 核心流程函数

| 函数名 | 位置 | 功能 |
|---|---|---|
| `Collect-All` | [L542-L658](file:///workspace/vcenter_inspect.ps1#L542-L658) | REST 数据采集主入口，14 步顺序执行 |
| `Collect-PowerCLI` | [L403-L535](file:///workspace/vcenter_inspect.ps1#L403-L535) | PowerCLI 回退采集入口，3 类数据 |
| `Eval-Findings` | [L663-L793](file:///workspace/vcenter_inspect.ps1#L663-L793) | 告警评估，返回 issues 列表 |
| `Render-Report` | [L798-L1856](file:///workspace/vcenter_inspect.ps1#L798-L1856) | HTML 报告渲染主函数 |
| `Render-Findings` | [L1741-L1749](file:///workspace/vcenter_inspect.ps1#L1741-L1749) | 单列 findings 渲染辅助函数 |

---

## 6. 数据流与依赖关系

### 6.1 主流程数据流

```
param() 输入参数
    ↓
Log-Banner (打印横幅)
    ↓
Collect-All (REST 采集)
    ├→ Login-VC → 设置 $Script:Session / $Script:ApiStyle
    ├→ Invoke-VC (N 次) → 写入 $Data.{各章节}
    └→ Logout-VC
    ↓
Collect-PowerCLI (可选，try/catch 隔离)
    ├→ Try-LoadPowerCLI → 检测模块
    ├→ Connect-PowerCLI → 连接 vCenter
    ├→ Get-VMHost → $Data.PCLI_Hosts
    ├→ Get-Snapshot → $Data.PCLI_Snapshots
    ├→ TriggeredAlarmState → $Data.PCLI_Alarms
    └→ Disconnect-PowerCLI
    ↓
Render-Report
    ├→ 预处理: 从 $Data 提取变量 + 统计聚合
    ├→ Eval-Findings → $findings 列表
    ├→ 构造 CSS (4 主题 + 业务样式)
    ├→ 构造 HTML (banner + cards + 19 sections + footer)
    └→ 注入 JS (scroll-spy + 主题切换)
    ↓
[System.IO.File]::WriteAllText → 输出 .html 文件
    ↓
exit 0
```

### 6.2 模块依赖图

```
vcenter_inspect.ps1
├→ [System.Net.HttpWebRequest]  —  .NET 内置，无外部依赖
├→ [System.Net.WebUtility]      —  HTML 编码
└→ (可选) VMware.VimAutomation.Core — PowerCLI 模块

html_to_docx.ps1
└→ Microsoft.Office.Interop.Word (COM) — 需要本机安装 Word

html_to_md.ps1
├→ [System.Text.RegularExpressions] —  .NET 内置
└→ [System.Net.WebUtility]           —  HTML 解码
```

---

## 7. 命令行参数参考

### 7.1 vcenter_inspect.ps1

| 参数 | 必需 | 默认值 | 类型 | 说明 |
|---|---|---|---|---|
| `-VCenter` | ✅ | - | string | vCenter IP 或 FQDN |
| `-Username` | ✅ | - | string | 用户名，推荐 `administrator@vsphere.local` |
| `-Password` | ✅ | - | string | 密码 |
| `-Output` | ❌ | `./report_<vc>_<date>.html` | string | 输出 HTML 路径 |
| `-ToolsSampleSize` | ❌ | `16` | int | VMware Tools 抽样数量（开机 VM） |
| `-SkipToolsSample` | ❌ | `$false` | switch | 跳过 Tools 抽样，节省 ~10 秒 |
| `-DebugDump` | ❌ | `$false` | switch | 每个 endpoint 的 HTTP code + 返回体写入 debug log |
| `-Quiet` | ❌ | `$false` | switch | 静默运行，不打印进度（CI / 计划任务用） |
| `-UsePowerCLI` | ❌ | 自动检测 | switch | 强制启用 PowerCLI 回退 |
| `-SkipPowerCLI` | ❌ | `$false` | switch | 强制跳过 PowerCLI 回退 |
| `-Theme` | ❌ | `'light'` | string | 报告默认主题: `light` / `dark` / `minimal` / `amber` |
| `-AccentColor` | ❌ | - | string | 单独覆盖主色调，hex 如 `#10b981` |

### 7.2 html_to_docx.ps1

| 参数 | 必需 | 默认值 | 说明 |
|---|---|---|---|
| `-InputPath` | ✅ | - | 输入 HTML 文件路径 |
| `-OutputPath` | ❌ | 同目录同名 `.docx` | 输出 DOCX 路径 |

### 7.3 html_to_md.ps1

| 参数 | 必需 | 默认值 | 说明 |
|---|---|---|---|
| `-InputPath` | ✅ | - | 输入 HTML 文件路径 |
| `-OutputPath` | ❌ | 同目录同名 `.md` | 输出 Markdown 路径 |

---

## 8. 运行方式

### 8.1 环境要求

- **跳板机**: Windows 10 / 11 / Server 2016+
- **PowerShell**: 5.1+（Windows 自带）或 PowerShell 7+
- **网络**: 能访问 vCenter 的 HTTPS 443 端口
- **账号**: vCenter Read-Only 角色即可

### 8.2 基本使用

```powershell
# 克隆仓库
git clone https://github.com/Aidan-996/VMware_vCenter_Inspect.git
cd VMware_vCenter_Inspect

# 执行巡检
.\vcenter_inspect.ps1 `
    -VCenter 10.0.0.20 `
    -Username administrator@vsphere.local `
    -Password 'YourPassword'

# 查看报告
# 生成文件: report_10.0.0.20_2026-05-25.html
```

### 8.3 格式转换

```powershell
# 转 Word（需安装 Office Word）
.\html_to_docx.ps1 -InputPath .\report_10.0.0.20_2026-05-25.html

# 转 Markdown
.\html_to_md.ps1 -InputPath .\report_10.0.0.20_2026-05-25.html
```

### 8.4 常用组合

```powershell
# 大规模 VM 环境，跳过 Tools 抽样
.\vcenter_inspect.ps1 -VCenter ... -SkipToolsSample

# 指定输出路径 + 静默模式（脚本化调用）
.\vcenter_inspect.ps1 -VCenter ... -Quiet -Output 'C:\reports\daily.html'

# Debug 模式（排查 API 问题）
.\vcenter_inspect.ps1 -VCenter ... -DebugDump

# 强制启用 PowerCLI 回退
.\vcenter_inspect.ps1 -VCenter ... -UsePowerCLI

# 切换主题 + 自定义主色
.\vcenter_inspect.ps1 -VCenter ... -Theme dark -AccentColor "#10b981"
```

### 8.5 自动化场景

#### Windows 计划任务（每周一巡检）

```powershell
$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument '-NoProfile -File C:\Tools\vcenter_inspect.ps1 -VCenter 10.0.0.20 -Username administrator@vsphere.local -Password "***" -Quiet -Output C:\Reports\weekly.html'

$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 8:00am

Register-ScheduledTask -TaskName 'vCenter Weekly Inspect' `
    -Action $action -Trigger $trigger -RunLevel Highest
```

#### 多 vCenter 批量巡检

```powershell
@(
    @{ vc='10.0.0.20'; out='vc01.html' }
    @{ vc='10.0.0.21'; out='vc02.html' }
    @{ vc='10.0.0.22'; out='vc03.html' }
) | ForEach-Object {
    .\vcenter_inspect.ps1 -VCenter $_.vc `
        -Username administrator@vsphere.local -Password '***' `
        -Output $_.out -Quiet
}
```

### 8.6 退出码

| Exit Code | 含义 |
|---|---|
| 0 | 成功完成 |
| 2 | 致命错误（登录失败等） |

---

## 9. 开发约定与踩坑记录

### 9.1 开发约定

1. **脚本编码**: 必须保存为 **UTF-8 + BOM**（PowerShell 5.1 兼容要求）
2. **零依赖**: 不引入新的外部模块（保持零依赖原则）
3. **章节命名**: 新增章节遵循 `Section-XX-<name>` 命名 + 同步 `Eval-Findings` 评估规则
4. **只读原则**: 不使用 POST/PUT/PATCH 修改类请求（DELETE 仅用于注销 session）
5. **错误降级**: API 失败时返回 `$null`，对应章节显示 N/A，不打断整体流程

### 9.2 PowerShell 5.1 踩坑记录

| 坑 | 现象 | 解法 |
|---|---|---|
| 脚本编码问题 | PS 5.1 默认按 ANSI/GBK 读 .ps1，中文乱码 | 脚本存为 **UTF-8 + BOM** |
| `Invoke-RestMethod` 编码 | 中文 VM 名 GBK 误解码 → `???` | 改用 `[System.Net.HttpWebRequest]` 手控 + UTF8 读流 |
| 自签证书拒绝 | vCenter 默认自签证书被拒绝 | 注入 `TrustAllCertsPolicy` 类 + 强制 TLS12 |
| 空数组计数 | `ConvertFrom-Json '[]'` 在 PS 5.1 返回 `$null`，`@($null).Count = 1` | 关键路径写 `if ($null -eq $x) { @() } else { @($x) }` |
| inline-if 解析 | 嵌套 inline-if 子表达式在 hashtable value 位置偶尔解析失败 | 预计算到变量，再引用 |
| UTF8 构造函数 | `[System.Text.UTF8Encoding]::new($false)` 在部分 PS 5.1 不可用 | 用 `New-Object System.Text.UTF8Encoding($false)` 替代 |

### 9.3 新增章节开发步骤

1. 在 `Collect-All` 中添加采集步骤（更新 `$Script:StepTotal`）
2. 数据存入 `$Data.<章节名>`
3. 在 `Eval-Findings` 中添加评估规则
4. 在 `Render-Report` 中添加 section（更新 `$toc` 数组）
5. 分配章节 ID（`sec-<name>`）和序号

---

## 10. 版本演进路线

### 10.1 已发布版本

| 版本 | 发布日期 | 核心内容 |
|---|---|---|
| **v1.0.0** | 2026-05-25 | 首发：17 章节 + dual-mode + 18 条 Findings + HTML/MD/DOCX |
| **v1.1.0** | 2026-06-19 | PowerCLI 回退层 + 快照/Alarm/Host 实时数据 + 10 条新规则 |
| **v1.2.0** | 2026-06-19 | 4 套主题系统 + 报告内实时切换 + localStorage 持久化 |

### 10.2 计划中（v1.3）

| 优先级 | 项目 | 说明 |
|---|---|---|
| P0 | 多 vCenter 批量 | `-VCenter @('vc1','vc2','vc3')` 一次跑一组 + 对比汇总 |
| P1 | Findings 基线对比 | 跟上次结果 diff，只输出新增 / 已解决告警 |
| P1 | 7.0 / 7.0U3 实测 | 路径已通过 dual-mode 匹配，缺真实环境验证 |
| P2 | `-RetryOnLoginFail` | 登录阶段也走指数退避（扛 sts-idmd 慢启动） |

### 10.3 设想中（v1.4 / v2.0）

- Telegram / 飞书 / 钉钉告警推送
- 配置文件化（阈值从 JSON 读取）
- 历史趋势图（SQLite + Chart.js）
- Health 整体评分（0-100 分）
- 国际化（英文报告模板）
- 多 vCenter Web 控制台（Flask / FastAPI dashboard）
- ESXi 直连模式
- vSAN 健康专题章节
- 容器化部署（Docker + helm）
- 报告 diff 可视化（并排对比）

---

## 附录: REST API 端点清单

### Appliance API

| 端点 | 方法 | 说明 |
|---|---|---|
| `/api/appliance/system/version` | GET | 系统版本信息 |
| `/api/appliance/system/uptime` | GET | 运行时间（秒） |
| `/api/appliance/system/time` | GET | 当前系统时间 |
| `/api/appliance/health/system` | GET | 系统健康状态 |
| `/api/appliance/health/storage` | GET | 存储健康状态 |
| `/api/appliance/health/mem` | GET | 内存健康状态 |
| `/api/appliance/health/swap` | GET | Swap 健康状态 |
| `/api/appliance/health/load` | GET | 负载健康状态 |
| `/api/appliance/health/database-storage` | GET | 数据库存储健康 |
| `/api/appliance/health/applmgmt` | GET | 应用管理健康 |
| `/api/appliance/health/software-packages` | GET | 软件包健康 |
| `/api/appliance/health/system/lastcheck` | GET | 上次健康检查时间 |
| `/api/appliance/networking` | GET | 网络配置（含 DNS） |
| `/api/appliance/networking/interfaces` | GET | 网卡列表 |
| `/api/appliance/ntp` | GET | NTP 服务器列表 |
| `/api/appliance/timesync` | GET | 时间同步模式 |
| `/api/appliance/access/ssh` | GET | SSH 状态 |
| `/api/appliance/access/shell` | GET | Bash Shell 状态 |
| `/api/appliance/access/dcui` | GET | DCUI 状态 |
| `/api/appliance/access/consolecli` | GET | Console CLI 状态 |
| `/api/appliance/recovery/backup/job` | GET | 备份任务历史 |
| `/api/appliance/recovery/backup/schedules` | GET | 备份计划 |
| `/api/appliance/services` | GET | 所有服务状态 |
| `/api/appliance/update` | GET | 更新状态 |

### vCenter API

| 端点 | 方法 | 说明 |
|---|---|---|
| `/api/session` | POST/DELETE | 登录 / 注销 |
| `/api/vcenter/deployment` | GET | 部署类型 |
| `/api/vcenter/datacenter` | GET | 数据中心列表 |
| `/api/vcenter/cluster` | GET | 集群列表 |
| `/api/vcenter/folder` | GET | 文件夹列表 |
| `/api/vcenter/resource-pool` | GET | 资源池列表 |
| `/api/vcenter/host` | GET | ESXi 主机列表 |
| `/api/vcenter/datastore` | GET | Datastore 列表 |
| `/api/vcenter/network` | GET | 网络 / Portgroup 列表 |
| `/api/vcenter/vm` | GET | 虚拟机列表 |
| `/api/vcenter/vm/{vm}/tools` | GET | 单 VM Tools 状态 |
| `/api/vcenter/certificate-management/vcenter/tls` | GET | vCenter TLS 证书 |

---

*本文档基于 v1.2.0 版本生成，最后更新: 2026-07-21*
