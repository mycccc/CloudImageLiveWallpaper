<#
.SYNOPSIS
    下载风云4B FY4B AGRI IMG REGI PCC GLL C013 云图，并配置桌面幻灯片、可注册计划任务。
.DESCRIPTION
    - 默认以脚本自身所在目录为保存目录（SaveDir）。
    - 每次运行都把桌面配置为幻灯片模式、相册文件夹指向 SaveDir、契合度填充、间隔 10 秒。
      相册文件夹通过 IDesktopWallpaper COM 接口设置（官方支持的方式）。
    - 从当前 UTC 向下对齐 15 分钟开始，追溯 HoursBack 小时（默认 8）。
    - 本地已存在跳过；404 视为“未发布”静默跳过。
    - 下载结束后清理超出时间窗口的旧图。
    - 日志按日期命名，自动清理非今日的日志。
    - 加 -InstallTask 时注册/覆盖计划任务 FY4B_CloudImage_Download，
      任务动作为 wscript.exe 调用同目录 2run_hidden.vbs（无窗口）。
.PARAMETER HoursBack
    往前追溯的小时数，默认 8。
.PARAMETER SaveDir
    保存目录。留空则使用脚本自身所在目录（$PSScriptRoot）。
.PARAMETER TimeoutSec
    单张图片下载超时秒数，默认 30。
.PARAMETER InstallTask
    传入后注册/更新计划任务（要求同目录存在 2run_hidden.vbs）。
#>

param(
    [int]$HoursBack = 8,
    [string]$SaveDir = '',
    [int]$TimeoutSec = 30,
    [switch]$InstallTask
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($SaveDir)) {
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $SaveDir = $PSScriptRoot
    } else {
        $SaveDir = (Get-Location).Path
    }
}
$SaveDir = $SaveDir.TrimEnd('\')

$TaskName   = 'FY4B_CloudImage_Download'
$TaskPath   = '\'
$BaseUrl    = 'https://img.nsmc.org.cn/CLOUDIMAGE/FY4B/AGRI/PCC/GLL/'
$FilePrefix = 'FY4B-_AGRI--_N_REGI_1050E_L2-_PCC-_C013_GLL_'
$FileSuffix = '_4000M_V0001.JPG'
$UserAgent  = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36'

if (-not (Test-Path -LiteralPath $SaveDir)) {
    New-Item -ItemType Directory -Path $SaveDir -Force | Out-Null
}

$todayStr = Get-Date -Format 'yyyyMMdd'
$LogFile  = Join-Path $SaveDir ('download_{0}.log' -f $todayStr)

Get-ChildItem -LiteralPath $SaveDir -Filter 'download_*.log' -File -ErrorAction SilentlyContinue | ForEach-Object {
    if ($_.Name -ne ('download_{0}.log' -f $todayStr)) {
        try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch { }
    }
}

function Write-Log {
    param([string]$Message)
    $line = '[' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] ' + $Message
    Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8
    Write-Host $line
}

Write-Log ('脚本路径：' + $PSCommandPath)
Write-Log ('保存目录：' + $SaveDir)
Write-Log ('回溯窗口：' + $HoursBack + ' 小时')

# ---------- 桌面幻灯片设置 ----------
function Set-DesktopSlideshow {
    param([string]$Folder)

    if (-not (Test-Path -LiteralPath $Folder)) {
        Write-Log ('桌面设置跳过：文件夹不存在 -> ' + $Folder)
        return
    }

    $images = Get-ChildItem -LiteralPath $Folder -Filter 'FY4B_*.jpg' -File -ErrorAction SilentlyContinue
    if (-not $images -or $images.Count -eq 0) {
        Write-Log '桌面设置跳过：文件夹中没有图片，等待下载完成后再设置'
        return
    }

    # 1. 基础注册表：填充、非平铺、幻灯片模式、10 秒间隔
    try {
        $deskPath = 'HKCU:\Control Panel\Desktop'
        Set-ItemProperty -Path $deskPath -Name 'WallpaperStyle' -Value 10 -Type String -Force
        Set-ItemProperty -Path $deskPath -Name 'TileWallpaper' -Value 0 -Type String -Force

        $wpPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers'
        if (-not (Test-Path -Path $wpPath)) { New-Item -Path $wpPath -Force | Out-Null }
        Set-ItemProperty -Path $wpPath -Name 'BackgroundType' -Value 2 -Type DWord -Force

        $ssPath = 'HKCU:\Control Panel\Personalization\Desktop Slideshow'
        if (-not (Test-Path -Path $ssPath)) { New-Item -Path $ssPath -Force | Out-Null }
        Set-ItemProperty -Path $ssPath -Name 'Interval' -Value 10000 -Type DWord -Force
    } catch {
        Write-Log ('基础注册表设置失败：' + $_.Exception.Message)
    }

    # 2. 通过 IDesktopWallpaper COM 接口设置幻灯片文件夹（官方方式）
    $slideshowSet = $false
    try {
        if (-not ('Fy4bWallpaperManager' -as [type])) {
            $comCode = @'
using System;
using System.Runtime.InteropServices;

public class Fy4bWallpaperManager
{
    [ComImport, Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD")]
    private class DesktopWallpaperClass { }

    [ComImport, Guid("B92B56A9-8B55-4E14-9A89-0199BBB6F93B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IDesktopWallpaper
    {
        void SetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID, [MarshalAs(UnmanagedType.LPWStr)] string wallpaper);
        [return: MarshalAs(UnmanagedType.LPWStr)]
        string GetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID);
        [return: MarshalAs(UnmanagedType.LPWStr)]
        string GetMonitorDevicePathAt(uint monitorIndex);
        uint GetMonitorDevicePathCount();
        void GetMonitorRECT([MarshalAs(UnmanagedType.LPWStr)] string monitorID, out RECT displayRect);
        void SetBackgroundColor(uint color);
        uint GetBackgroundColor();
        void SetPosition(int position);
        int GetPosition();
        void SetSlideshow([MarshalAs(UnmanagedType.Interface)] object items);
        [return: MarshalAs(UnmanagedType.Interface)]
        object GetSlideshow();
        void SetSlideshowOptions(int options, uint slideshowTick);
        void GetSlideshowOptions(out int options, out uint slideshowTick);
        void AdvanceSlideshow([MarshalAs(UnmanagedType.LPWStr)] string monitorID, int direction);
        void Enable(bool enable);
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHCreateItemFromParsingName(
        [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
        IntPtr pbc,
        ref Guid riid,
        out IntPtr ppv);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHCreateShellItemArrayFromShellItem(
        IntPtr psi,
        ref Guid riid,
        out IntPtr ppsiItemArray);

    public static int SetSlideshowFolder(string folderPath, uint intervalMs)
    {
        Guid iidShellItem = new Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE");
        Guid iidShellItemArray = new Guid("B63EA76D-1F85-456F-A19C-48159EFA858B");

        IntPtr shellItemPtr;
        int hr = SHCreateItemFromParsingName(folderPath, IntPtr.Zero, ref iidShellItem, out shellItemPtr);
        if (hr != 0) return hr;

        try
        {
            IntPtr shellItemArrayPtr;
            hr = SHCreateShellItemArrayFromShellItem(shellItemPtr, ref iidShellItemArray, out shellItemArrayPtr);
            if (hr != 0) return hr;

            try
            {
                var wp = (IDesktopWallpaper)new DesktopWallpaperClass();
                wp.SetSlideshow(Marshal.GetObjectForIUnknown(shellItemArrayPtr));
                wp.SetSlideshowOptions(0, intervalMs);
                return 0;
            }
            finally
            {
                Marshal.Release(shellItemArrayPtr);
            }
        }
        finally
        {
            Marshal.Release(shellItemPtr);
        }
    }
}
'@
            Add-Type -TypeDefinition $comCode -ErrorAction Stop
        }

        $rc = [Fy4bWallpaperManager]::SetSlideshowFolder($Folder, [uint32]10000)
        if ($rc -eq 0) {
            Write-Log ('幻灯片文件夹已设置：' + $Folder + '（IDesktopWallpaper）')
            $slideshowSet = $true
        } else {
            Write-Log ('IDesktopWallpaper.SetSlideshow 失败，HRESULT=0x' + $rc.ToString('X8'))
        }
    } catch {
        Write-Log ('IDesktopWallpaper 调用失败：' + $_.Exception.Message)
    }

    # 3. 回退：若 COM 失败，用最新一张图片做壁纸，避免纯色
    if (-not $slideshowSet) {
        try {
            $sig = @'
using System;
using System.Runtime.InteropServices;
public class PInvokeSlideshow {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
'@
            if (-not ('PInvokeSlideshow' -as [type])) {
                Add-Type -TypeDefinition $sig -ErrorAction Stop
            }
            $latest = $images | Sort-Object Name -Descending | Select-Object -First 1
            $rc2 = [PInvokeSlideshow]::SystemParametersInfo(0x0014, 0, $latest.FullName, 0x03)
            Write-Log ('回退：已设置最新图片为壁纸 ' + $latest.Name + '（SPI 返回 ' + $rc2 + '）')
            # 再确保一次幻灯片模式（SPI 可能把它改回图片模式）
            Set-ItemProperty -Path $wpPath -Name 'BackgroundType' -Value 2 -Type DWord -Force
        } catch {
            Write-Log ('回退设置壁纸失败：' + $_.Exception.Message)
        }
    }

    try {
        Start-Process -FilePath 'RUNDLL32.EXE' -ArgumentList 'USER32.DLL,UpdatePerUserSystemParameters' -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue
    } catch { }
}

# ---------- 可选：注册/更新计划任务 ----------
if ($InstallTask) {
    $scriptPath = $PSCommandPath
    if ([string]::IsNullOrWhiteSpace($scriptPath)) { $scriptPath = $MyInvocation.MyCommand.Path }
    if ([string]::IsNullOrWhiteSpace($scriptPath)) { throw '无法确定脚本自身路径，无法安装计划任务。' }

    $scriptDir = Split-Path -Parent $scriptPath
    $vbsPath   = Join-Path $scriptDir '2run_hidden.vbs'
    if (-not (Test-Path -LiteralPath $vbsPath)) {
        throw ('缺少无窗口启动器：' + $vbsPath + '。请先运行 1install.cmd 生成它。')
    }

    $action = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument ('"' + $vbsPath + '"')

    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
        -RepetitionInterval (New-TimeSpan -Minutes 15) `
        -RepetitionDuration (New-TimeSpan -Days 3650)

    $settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Limited

    Register-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath `
        -Action $action -Trigger $trigger -Settings $settings -Principal $principal `
        -Force -ErrorAction Stop | Out-Null

    Write-Log ('计划任务已注册：' + $TaskPath + $TaskName +
               '（wscript.exe 调用 ' + $vbsPath + '，用户 ' + $currentUser + '，每 15 分钟）')
}

# ---------- 每次运行都确保桌面幻灯片指向 SaveDir ----------
Set-DesktopSlideshow -Folder $SaveDir

# ---------- 计算时间点 ----------
$nowUtc = [DateTime]::UtcNow
$alignedMinute = [math]::Floor($nowUtc.Minute / 15) * 15
$latestPoint = $nowUtc.Date.AddHours($nowUtc.Hour).AddMinutes($alignedMinute)
$oldestPoint = $latestPoint.AddHours(-$HoursBack)

$startMsg = '开始运行：最新 UTC {0:yyyy-MM-dd HH:mm}，最早 UTC {1:yyyy-MM-dd HH:mm}，窗口 {2} 小时' -f $latestPoint, $oldestPoint, $HoursBack
Write-Log $startMsg

$points = New-Object System.Collections.Generic.List[DateTime]
$t = $latestPoint
while ($t -ge $oldestPoint) {
    $points.Add($t)
    $t = $t.AddMinutes(-15)
}

$okCount       = 0
$skipCount     = 0
$notFoundCount = 0
$failCount     = 0

foreach ($p in $points) {
    $localName = 'FY4B_{0:yyyyMMdd_HHmm}.jpg' -f $p
    $localPath = Join-Path $SaveDir $localName

    if (Test-Path -LiteralPath $localPath) {
        $skipCount++
        continue
    }

    $endTime    = $p.AddSeconds(899)
    $startStr   = $p.ToString('yyyyMMddHHmmss')
    $endStr     = $endTime.ToString('yyyyMMddHHmmss')
    $remoteName = $FilePrefix + $startStr + '_' + $endStr + $FileSuffix
    $url        = $BaseUrl + $remoteName

    try {
        Invoke-WebRequest -Uri $url -OutFile $localPath -TimeoutSec $TimeoutSec `
            -UseBasicParsing -UserAgent $UserAgent -ErrorAction Stop | Out-Null
        $okCount++
        Write-Log ('下载成功：' + $localName)
    }
    catch {
        if (Test-Path -LiteralPath $localPath) {
            try { Remove-Item -LiteralPath $localPath -Force } catch { }
        }
        $status = $null
        if ($null -ne $_.Exception.Response) {
            $status = [int]$_.Exception.Response.StatusCode
        }
        if ($status -eq 404) {
            $notFoundCount++
        } else {
            $failCount++
            Write-Log ('下载失败：' + $localName + ' (' + $_.Exception.Message + ')')
        }
    }
}

$deleted = 0
$files = Get-ChildItem -LiteralPath $SaveDir -Filter 'FY4B_*.jpg' -File
foreach ($f in $files) {
    $m = [regex]::Match($f.Name, '^FY4B_(\d{8}_\d{4})\.jpg$')
    if ($m.Success) {
        try {
            $fileTime = [DateTime]::ParseExact($m.Groups[1].Value, 'yyyyMMdd_HHmm', [Globalization.CultureInfo]::InvariantCulture)
            if ($fileTime -lt $oldestPoint) {
                Remove-Item -LiteralPath $f.FullName -Force
                $deleted++
            }
        } catch { }
    }
}

$doneMsg = '完成：新增 ' + $okCount + '，跳过 ' + $skipCount + '，未发布 ' + $notFoundCount + '，失败 ' + $failCount + '，清理 ' + $deleted
Write-Log $doneMsg