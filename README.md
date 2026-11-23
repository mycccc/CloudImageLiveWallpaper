# CloudImageLiveWallpaper
```markdown
# CloudImageLiveWallpaper

把风云四号B星（FY-4B）最新云图自动下载并设置为 Windows 桌面幻灯片壁纸。

## 简介

[国家卫星气象中心（NSMC）](https://img.nsmc.org.cn/) 持续发布 FY-4B 的风云云图。本项目通过 Windows 计划任务每 15 分钟自动拉取最近若干小时的云图，保存到本地文件夹，并把该文件夹配置为 Windows 桌面幻灯片来源，让桌面以 10 秒间隔轮播云图演变。

全部逻辑由三个文件实现，无需安装任何第三方软件，不修改系统文件，卸载只需删任务、删文件夹。

## 特性

- **自动下载**：每 15 分钟同步一次最近 N 小时（默认 8 小时）的 FY-4B AGRI 云图，按 15 分钟间隔逐张获取。
- **断点续跑**：本地已存在的图片自动跳过，网络失败不影响已有序列。
- **自动清理**：超出时间窗口的旧图自动删除，文件夹大小稳定可控。
- **桌面幻灯片**：通过官方 `IDesktopWallpaper` COM 接口把下载目录设为幻灯片相册，契合度固定为「填充」，间隔为 Windows 允许的 10 秒下限。
- **持续自修复**：每次运行都会重写幻灯片相关注册表，防止被其他软件或手动操作改回。
- **无窗口运行**：计划任务通过 `wscript.exe` 调用 VBS 启动器，全程无 PowerShell 窗口闪烁。
- **启用/禁用开关**：同一份安装脚本可一键启用或禁用计划任务，任务定义保留，随时恢复。
- **按日期日志**：每日一个日志文件，自动清理历史日志。

## 系统要求

- Windows 10 / Windows 11
- PowerShell 5.1 及以上（系统自带）
- 首次部署需管理员权限（注册计划任务）

## 目录结构

```
CloudImageLiveWallpaper\
├── 1install.cmd          # 首次部署 / 启用 / 禁用菜单
├── 2run_hidden.vbs       # 无窗口启动器（由 1install.cmd 生成/覆盖）
├── 3download_fy4b.ps1    # 业务主脚本：下载、桌面配置、清理
├── FY4B_YYYYMMDD_HHmm.jpg    # 下载产生的云图（自动管理）
└── download_YYYYMMDD.log     # 当日日志（自动管理）
```

## 快速开始

1. 下载本仓库的 `1install.cmd` 和 `3download_fy4b.ps1`，放在同一个可写目录（例如 `C:\Users\<用户名>\Pictures\liveFY4B\`）。
2. 右键 `1install.cmd`，选择 **以管理员身份运行**。
3. 在菜单中选择 `[1] Enable`，按提示输入回溯小时数（直接回车使用默认 8 小时）。
4. 脚本会依次：
   - 生成 `2run_hidden.vbs`；
   - 注册计划任务 `FY4B_CloudImage_Download`（每 15 分钟触发）；
   - 立即执行一次完整流程：配置桌面幻灯片、下载历史云图、清理超窗旧图。
5. 完成后，桌面开始以 10 秒间隔轮播云图。

## 启用 / 禁用

再次运行 `1install.cmd`：

| 选项 | 作用 |
| :--- | :--- |
| `[1] Enable` | 生成/覆盖 VBS 启动器，注册或覆盖计划任务，立即下载一次。适合首次部署、修改回溯小时数、从禁用状态恢复。 |
| `[2] Disable` | 禁用计划任务（`schtasks /change /disable`）。任务定义、触发器、脚本全部保留，下次选 1 即可恢复。 |
| `[3] Exit` | 直接退出。 |

> 禁用只是停掉自动下载，不会删除本地已下载的图片，也不会撤销桌面幻灯片设置。若想彻底移除，见下文「卸载」。

## 修改回溯小时数

- **通过安装脚本**：重新运行 `1install.cmd` → 选 `[1] Enable` → 输入新的小时数。脚本会覆盖 `2run_hidden.vbs` 和计划任务，并在当次下载后立即按新窗口清理旧图。
- **临时手动修改**：直接编辑 `2run_hidden.vbs` 中的 `-HoursBack N`，下次计划任务触发时生效。但下次运行 `1install.cmd` 会被覆盖回菜单输入值，长期修改建议走安装脚本。

## 配置参数

`3download_fy4b.ps1` 顶部 `param` 块可调整：

| 参数 | 默认值 | 说明 |
| :--- | :--- | :--- |
| `-HoursBack` | `8` | 回溯小时数 |
| `-SaveDir` | 脚本所在目录 | 图片、日志保存位置 |
| `-TimeoutSec` | `30` | 单张图片下载超时（秒） |
| `-InstallTask` | 关闭 | 传入时注册/覆盖计划任务 |

## 工作原理

### 数据源

FY-4B AGRI 云图 URL 命名规律稳定：

```
https://img.nsmc.org.cn/CLOUDIMAGE/FY4B/AGRI/PCC/GLL/
FY4B-_AGRI--_N_REGI_1050E_L2-_PCC-_C013_GLL_
{开始时间yyyyMMddHHmmss}_{结束时间yyyyMMddHHmmss}_4000M_V0001.JPG
```

- 开始时间为 UTC，15 分钟对齐；
- 结束时间 = 开始时间 + 14 分 59 秒；
- 产品发布通常滞后当前时间 45～60 分钟，脚本不依赖这一滞后值，而是从当前 UTC 向下对齐 15 分钟开始枚举，未发布的时次返回 404 时静默跳过。

### 执行流程（每次计划任务触发）

```
计划任务 → wscript.exe → 2run_hidden.vbs → powershell.exe (隐藏)
    → 3download_fy4b.ps1
        ├─ 写当日日志，清理非今日日志
        ├─ （仅 -InstallTask 时）注册/覆盖计划任务
        ├─ 配置桌面：BackgroundType=2, WallpaperStyle=10,
        │            Interval=10000, IDesktopWallpaper.SetSlideshow(SaveDir)
        ├─ 枚举 [当前 UTC-15min对齐 - HoursBack, 当前 UTC-15min对齐]
        │   逐张构造 URL → 本地存在则跳过 → 不存在则下载 → 404 跳过
        └─ 删除早于窗口起点的本地图片
```

### 为什么需要 VBS 启动器

`powershell.exe` 是控制台程序，任务计划程序启动它时 Windows 会先分配控制台窗口，即使加 `-WindowStyle Hidden` 仍会出现瞬间闪烁。`wscript.exe` 是 GUI 子系统程序，本身无控制台；它通过 `WScript.Shell.Run(cmd, 0, False)` 的 `0`（隐藏窗口）标志启动 PowerShell，在进程创建阶段就抑制窗口，实现全程无闪烁。

### 为什么用 IDesktopWallpaper 而不是直接写注册表

`HKCU\...\Explorer\Wallpapers\SlideshowDirectoryPath1` 存储的路径经过加密，无法通过脚本写入有效值。`IDesktopWallpaper::SetSlideshow(IShellItemArray)` 是微软文档化的官方接口，可以直接把文件夹设为幻灯片来源。

### 关于 10 秒间隔

Windows 10/11 原生幻灯片的最小间隔是 **10 秒（10000 毫秒）**，写入更小的值会被系统覆盖回 10000。这不是本项目的限制，而是系统行为。实测即使在注册表中写入 1000，实际切换仍为 10 秒。如需更快切换，只能绕开系统幻灯片机制（第三方动态壁纸工具或视频壁纸方案）。

## 已知限制

- **10 秒切换下限**：Windows 原生幻灯片的硬下限，脚本无法突破。
- **发布延迟**：最新云图通常滞后当前时间 45～60 分钟，脚本不追求「实时最新」，窗口按固定间隔枚举。
- **URL 规则依赖**：脚本直接构造 URL，假设 NSMC 的路径、文件名结构、通道标识（`C013`）、分辨率（`4000M`）、版本号（`V0001`）保持不变。若官方调整，需要相应修改脚本中的 `BaseUrl`、`FilePrefix`、`FileSuffix`。
- **单产品**：当前仅支持 `FY4B AGRI IMG REGI PCC GLL C013`。若要下载其他通道（如 `C012` 可见光）或其他产品，需改参数并考虑多目录管理。
- **Interactive 登录态**：计划任务以当前用户 Interactive 运行，仅在用户登录时触发。无人登录时桌面本身不可见，这一行为符合使用场景。

## 卸载

1. 运行 `1install.cmd` → 选 `[2] Disable` 禁用计划任务；
2. 打开任务计划程序，删除根路径下的 `FY4B_CloudImage_Download`；
3. 在「设置 → 个性化 → 背景」中把背景改回图片或纯色；
4. 删除整个 `liveFY4B` 文件夹即可。

## 常见问题

**任务计划程序里有多条同名任务**

本项目固定使用 `TaskPath=\` + `TaskName=FY4B_CloudImage_Download` + `-Force` 覆盖注册。无论从哪个目录运行 `1install.cmd`，都只会保留一条。若发现重复，删除多余的并重跑 `1install.cmd` 即可。

**移动了脚本目录后计划任务不工作**

计划任务参数中写入了脚本的绝对路径。移动目录后需要重新运行 `1install.cmd` 选 `[1] Enable`，让任务指向新路径。

**桌面没有轮播**

依次检查：任务计划程序中任务状态是否为「已启用」；日志是否有下载成功记录；「设置 → 个性化 → 背景」中类型是否为「幻灯片放映」、相册是否指向脚本目录；文件夹里是否至少有 2 张图片。

**计划任务运行了但窗口仍然闪一下**

检查计划任务的动作是否为 `wscript.exe` + `2run_hidden.vbs`。如果之前手动注册过直接调用 `powershell.exe` 的任务，需要重新运行 `1install.cmd` 选 `[1] Enable` 覆盖为新动作。

## 致谢与免责

- 数据来源：[国家卫星气象中心](https://www.nsmc.org.cn/)，请遵守其数据使用条款。
- 本项目仅做下载与本地桌面展示，不缓存、不分发、不修改原始云图。
- 脚本按现状提供，不保证长期兼容 NSMC 的 URL 规则变化。

## License

MIT
```
