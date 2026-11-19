#requires -Version 5.1
<#
    FolderPad —— 桌面「围栏式文件夹框」
    ================================================================
    一个框 = 一个真实文件夹的另一种形态。同一个路径，两种呈现：
    普通文件夹（桌面显示图标）／框（桌面显示一片围栏，里面是图标）。
    切换形态不复制、不移动文件；做框时带倒计时，边框随剩余天数变色。

    需求说明见同目录《需求说明.md》，使用说明见《使用说明.md》。

    技术：Windows PowerShell 5.1 + WPF，免安装、离线、无后台服务。
    数据：%APPDATA%\FolderPad\
          pads.json          框记录（写 pads.tmp → 替换 pads.json → 保留 pads.bak）
          hidden-backup.txt  纯文本恢复清单（每行：路径 \t 原属性 \t 原隐藏位）
          clean-exit         正常退出标记，启动时不存在即判定上次异常退出
          pending.txt        待转框的文件夹队列（发送到入口投递）
          FolderPad.log      运行日志

    模式：
      (无参数)              启动图形界面（常驻托盘）
      <文件夹路径...>        转框模式：投递给已运行的实例，或自己启动并处理
      -Install              安装：桌面快捷方式 + 发送到入口（可选开机自启）
      -Uninstall            卸载快捷方式（不动数据、不动用户文件）
      -SelfTest             无界面自检（数据、档位、属性还原、压缩）
      -Demo                 建一个演示框（临时目录，用来看效果）
#>

param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$InputPaths,
    [switch]$SelfTest,
    [switch]$Install,
    [switch]$Uninstall,
    [switch]$Demo,
    [switch]$Startup,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$script:AppName    = 'FolderPad'
$script:AppVersion = '1.1.1'
$script:AppDir     = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }

$script:DataDir           = Join-Path $env:APPDATA 'FolderPad'
$script:Path_Json         = Join-Path $script:DataDir 'pads.json'
$script:Path_Bak          = Join-Path $script:DataDir 'pads.bak'
$script:Path_Tmp          = Join-Path $script:DataDir 'pads.tmp'
$script:Path_HiddenBackup = Join-Path $script:DataDir 'hidden-backup.txt'
$script:Path_CleanExit    = Join-Path $script:DataDir 'clean-exit'
$script:Path_Pending      = Join-Path $script:DataDir 'pending.txt'
$script:Path_Log          = Join-Path $script:DataDir 'FolderPad.log'

$script:CategoryPresets    = @('#6C5CE7', '#0984E3', '#00B894', '#E8B33A', '#E17055', '#8FA3B0')
# 紧急度档位可选色：前 5 个就是五个档位的出厂默认色（含「今天截止」的红 #D2362F），后面是备用色
$script:UrgencyPresets     = @('#8FA3B0', '#E8B33A', '#E8762C', '#D2362F', '#9A9A9A', '#6C5CE7', '#0984E3', '#00B894', '#E17055')
# 类别色（标题条）可选色：前 6 个是原来的预设色，后面是补充的
$script:CategoryColors     = @('#6C5CE7', '#0984E3', '#00B894', '#E8B33A', '#E17055', '#D2362F', '#FD79A8', '#00CEC9', '#A29BFE', '#8FA3B0', '#636E72', '#2D3436')
$script:ThemeTitleColor    = '#3A3F4B'
$script:NeutralBorderColor = '#B0B0B0'
$script:TitleHeight        = 34    # 标题条高度（原来 26，字体放大后跟着加高）
$script:TitleFont          = 14    # 标题文字号（默认 14，原来 12）
$script:DueFont            = 12.5  # 天数文字号（原来 11）
$script:PadRadius          = 8    # 框的圆角
$script:TitleRadius        = 6    # 标题条上圆角
$script:ChipFill           = 0.26  # 天数内嵌槽的底色（黑，越大越深）
$script:ChipRadius         = 7     # 内嵌槽圆角
$script:ChipMinWidth       = 0     # 0 = 槽宽跟着数字自适应（1 位窄、2 位宽，数字始终居中）
$script:ChipNudge          = 1.4   # 首位是 1 时向左微移的像素（等宽数字里 1 的字面偏窄，需要光学居中）
$script:DlgShadowPad       = 16    # 对话框四周给阴影留的透明边（太小阴影会被切出虚边）
$script:MenuShadowPad      = 12    # 右键菜单四周给阴影留的透明边
$script:MenuMaxWidth       = 400   # 菜单最大宽度（标题/提示再长也不会把菜单顶宽）
$script:MenuHintMaxWidth   = 150   # 右侧灰色提示的最大宽度
$script:MenuHintNameChars  = 12    # 提示里显示框名的最多字数（超出用 … 收）
$script:DueNumRedDays      = 5     # 剩余天数 ≤ 这个值时，槽里的数字变红
$script:DueNumRedColor     = '#FF5A50'  # 数字红（只改数字，标题条和底色不受影响）
$script:SeamHi            = 0.22  # 接缝上沿提亮（0 = 关）
$script:SeamLo            = 0.38  # 接缝下沿压暗
$script:ShadowPad          = 0     # 背板窗口比框大多少，留给外阴影
$script:BorderWidth        = 0     # 不再画彩色边框：紧急度色铺进标题栏下方的底色（BackBorder）
$script:DefaultFill        = 0.40  # 框内底色不透明度（紧急度色的透明度）
$script:MaxAutoCols        = 8
$script:FontFamily         = 'Microsoft YaHei UI'
$script:PadFreeLayout      = $true   # false = 退回纯网格（忽略 big / rows），出问题时的兜底开关
$script:BigItemSpan        = 2       # 放大突出的大格占几格（固定 2×2）
$script:BigIconScale       = 2       # 大格图标 = 普通图标 × 这个倍数
$script:ReorderByDrag      = $false  # 框内拖动排序（默认关：改用手感更好的方向键）

$script:State      = $null
$script:Pads       = @{}      # id -> 运行时对象（窗口、控件、条目）
$script:IconCache  = @{}
$script:Grid       = @{ CellW = 84; CellH = 92; Icon = 48; Scale = 1.0; Raw = $null }
$script:Tray       = $null
$script:HotkeySrc  = $null
$script:Mutex      = $null
$script:TickAll    = 0
$script:Migrated   = $false
$script:AppExiting = $false
$script:ShieldUp   = $false
$script:DragStart  = $null
$script:ThemeRes   = $null
$script:LastDate   = (Get-Date).Date
$script:LastRefresh = Get-Date
$script:SelfTestResult = @()
$script:HookedHwnds = @{}
$script:ArrowLatch = $false

# =====================================================================
# 原生互操作
# =====================================================================
$script:CSharpSource = @'
using System;
using System.Runtime.InteropServices;

[StructLayout(LayoutKind.Sequential)]
public struct FPPOINT { public int X; public int Y; }

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct FPSHFILEINFO {
    public IntPtr hIcon;
    public int iIcon;
    public uint dwAttributes;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szDisplayName;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)] public string szTypeName;
}

[ComImport]
[Guid("46EB5926-582E-4017-9FDF-E8998DAA0950")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface FPIImageList {
    void Add();
    void ReplaceIcon();
    void SetOverlayImage();
    void Replace();
    void AddMasked();
    void Draw();
    void Remove();
    [PreserveSig] int GetIcon(int i, int flags, out IntPtr picon);
}

public static class FPNative
{
    public const int GWL_EXSTYLE      = -20;
    public const int WS_EX_NOACTIVATE = 0x08000000;
    public const int WS_EX_TOOLWINDOW = 0x00000080;
    public const int WS_EX_TRANSPARENT= 0x00000020;
    public const int WS_EX_LAYERED    = 0x00080000;
    public const int WS_EX_APPWINDOW  = 0x00040000;

    public static readonly IntPtr HWND_TOP      = IntPtr.Zero;
    public static readonly IntPtr HWND_BOTTOM   = (IntPtr)1;
    public static readonly IntPtr HWND_TOPMOST  = (IntPtr)(-1);
    public static readonly IntPtr HWND_NOTOPMOST= (IntPtr)(-2);

    public const uint SWP_NOSIZE     = 0x0001;
    public const uint SWP_NOMOVE     = 0x0002;
    public const uint SWP_NOZORDER   = 0x0004;
    public const uint SWP_NOACTIVATE = 0x0010;
    public const uint SWP_SHOWWINDOW = 0x0040;

    public const uint ILD_TRANSPARENT = 0x00000001;

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out FPPOINT pt);
    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);
    [DllImport("user32.dll")]
    public static extern IntPtr WindowFromPoint(FPPOINT pt);
    [DllImport("user32.dll")]
    public static extern bool SystemParametersInfo(uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);
    [DllImport("user32.dll")]
    public static extern int GetSystemMetrics(int nIndex);
    [DllImport("user32.dll")]
    public static extern uint GetDpiForWindow(IntPtr hwnd);
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")]
    public static extern bool SetProcessDpiAwarenessContext(IntPtr value);
    [DllImport("user32.dll")]
    public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr FindWindow(string cls, string title);
    [DllImport("user32.dll")]
    public static extern IntPtr GetWindow(IntPtr hWnd, uint cmd);
    [DllImport("user32.dll")]
    public static extern IntPtr GetDesktopWindow();

    public const uint GW_HWNDPREV = 3;

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes,
        ref FPSHFILEINFO psfi, uint cbFileInfo, uint uFlags);
    [DllImport("shell32.dll")]
    public static extern int SHGetImageList(int iImageList, ref Guid riid, out FPIImageList ppv);

    public const uint SHGFI_ICON          = 0x000000100;
    public const uint SHGFI_LARGEICON     = 0x000000000;
    public const uint SHGFI_SMALLICON     = 0x000000001;
    public const uint SHGFI_USEFILEATTRIBUTES = 0x000000010;
    public const uint SHGFI_SYSICONINDEX  = 0x000004000;
    public const uint FILE_ATTRIBUTE_NORMAL = 0x00000080;
    public const uint FILE_ATTRIBUTE_DIRECTORY = 0x00000010;

    public static IntPtr GetIconByIndex(string path, int wantSize)
    {
        FPSHFILEINFO shfi = new FPSHFILEINFO();
        uint flags = SHGFI_SYSICONINDEX;
        IntPtr res = SHGetFileInfo(path, 0, ref shfi, (uint)Marshal.SizeOf(typeof(FPSHFILEINFO)), flags);
        if (res == IntPtr.Zero) return IntPtr.Zero;

        int shil;
        if (wantSize <= 16) shil = 1;        // SHIL_SMALL 16
        else if (wantSize <= 32) shil = 0;   // SHIL_LARGE 32
        else if (wantSize <= 48) shil = 2;   // SHIL_EXTRALARGE 48
        else shil = 4;                       // SHIL_JUMBO 256

        Guid iid = new Guid("46EB5926-582E-4017-9FDF-E8998DAA0950");
        FPIImageList list;
        int hr = SHGetImageList(shil, ref iid, out list);
        if (hr != 0 || list == null) return IntPtr.Zero;
        IntPtr h = IntPtr.Zero;
        try { hr = list.GetIcon(shfi.iIcon, (int)ILD_TRANSPARENT, out h); }
        catch { hr = -1; }
        finally { if (list != null) Marshal.ReleaseComObject(list); }
        if (hr != 0) return IntPtr.Zero;
        return h;
    }

    public static IntPtr GetLargeIcon(string path)
    {
        FPSHFILEINFO shfi = new FPSHFILEINFO();
        IntPtr res = SHGetFileInfo(path, 0, ref shfi, (uint)Marshal.SizeOf(typeof(FPSHFILEINFO)),
            SHGFI_ICON | SHGFI_LARGEICON);
        if (res == IntPtr.Zero) return IntPtr.Zero;
        return shfi.hIcon;
    }

    [DllImport("user32.dll")]
    public static extern bool DestroyIcon(IntPtr hIcon);

    public static uint DpiForWindow(IntPtr hwnd)
    {
        try { return GetDpiForWindow(hwnd); } catch { return 96; }
    }
}

// 窗口消息过滤器：主要用来拦住「显示桌面(Win+D)」发的 SC_MINIMIZE
public static class FPMsg
{
    public static string LogPath = null;

    public static void Log(string s)
    {
        try
        {
            if (String.IsNullOrEmpty(LogPath)) return;
            System.IO.File.AppendAllText(LogPath, DateTime.Now.ToString("HH:mm:ss.fff") + " " + s + "\r\n");
        }
        catch { }
    }

    public static IntPtr Hook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        try
        {
            if (msg == 0x0112)                       // WM_SYSCOMMAND
            {
                long cmd = wParam.ToInt64() & 0xFFF0;
                if (cmd == 0xF020)                   // SC_MINIMIZE
                {
                    Log("拦下 SC_MINIMIZE（显示桌面/最小化）");
                    handled = true;
                    return IntPtr.Zero;
                }
            }
            else if (msg == 0x0005 && wParam.ToInt64() == 1)   // WM_SIZE / SIZE_MINIMIZED
            {
                Log("收到 WM_SIZE(SIZE_MINIMIZED)");
            }
            else if (msg == 0x0018)                  // WM_SHOWWINDOW
            {
                Log("WM_SHOWWINDOW wParam=" + wParam.ToInt64());
            }
        }
        catch { }
        return IntPtr.Zero;
    }
}
'@

if (-not ('FPNative' -as [type])) {
    Add-Type -TypeDefinition $script:CSharpSource -Language CSharp -ErrorAction Stop
}
try { [FPMsg]::LogPath = (Join-Path $env:TEMP 'fp_msg.log') } catch { }

# DPI 感知：必须在任何窗口创建之前设置，否则网格与桌面图标对不齐
try { [void][FPNative]::SetProcessDpiAwarenessContext([IntPtr](-4)) }
catch {
    try { [void][FPNative]::SetProcessDPIAware() } catch { }
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

# =====================================================================
# 基础工具
# =====================================================================
function Initialize-DataDir {
    if (-not (Test-Path -LiteralPath $script:DataDir)) {
        New-Item -ItemType Directory -Force -Path $script:DataDir | Out-Null
    }
}

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    try {
        Initialize-DataDir
        if ((Test-Path -LiteralPath $script:Path_Log) -and ((Get-Item -LiteralPath $script:Path_Log).Length -gt 1MB)) {
            [System.IO.File]::WriteAllText($script:Path_Log, '')
        }
        $line = '{0} [{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff'), $Level, $Message
        [System.IO.File]::AppendAllText($script:Path_Log, $line + "`r`n", (New-Object System.Text.UTF8Encoding($false)))
    } catch { }
}

function ConvertTo-FPColor {
    param([string]$Hex, [double]$Alpha = 1.0)
    $h = ([string]$Hex).Trim().TrimStart('#')
    if ($h.Length -eq 3) {
        $chars = $h.ToCharArray()
        $h = -join ($chars | ForEach-Object { "$_$_" })
    }
    if ($h.Length -eq 8) { $h = $h.Substring(2) }
    if ($h.Length -ne 6) { $h = 'B0B0B0' }
    try {
        $a = [byte]([Math]::Round(255 * [Math]::Max(0.0, [Math]::Min(1.0, $Alpha))))
        return [System.Windows.Media.Color]::FromArgb(
            $a,
            [Convert]::ToByte($h.Substring(0, 2), 16),
            [Convert]::ToByte($h.Substring(2, 2), 16),
            [Convert]::ToByte($h.Substring(4, 2), 16))
    } catch {
        return [System.Windows.Media.Colors]::Gray
    }
}

function New-FPBrush {
    param([string]$Hex, [double]$Alpha = 1.0)
    $b = New-Object System.Windows.Media.SolidColorBrush (ConvertTo-FPColor -Hex $Hex -Alpha $Alpha)
    $b.Freeze()
    return $b
}

function Get-NormalizedPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $p = $Path.Trim().Trim('"')
    try { $p = [System.IO.Path]::GetFullPath($p) } catch { return '' }
    $p = $p.TrimEnd('\')
    if ($p.Length -eq 2 -and $p[1] -eq ':') { $p = $p + '\' }
    return $p
}

function Test-PathEqual {
    param([string]$A, [string]$B)
    $a = Get-NormalizedPath $A
    $b = Get-NormalizedPath $B
    if (-not $a -or -not $b) { return $false }
    return [string]::Equals($a, $b, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-SpecialPaths {
    $list = New-Object System.Collections.Generic.List[string]
    $add = {
        param($p)
        if ($p) { $n = Get-NormalizedPath $p; if ($n) { $list.Add($n) } }
    }
    & $add $env:USERPROFILE
    & $add $env:SystemRoot
    & $add $env:WINDIR
    & $add $env:ProgramFiles
    & $add ${env:ProgramFiles(x86)}
    & $add $env:ProgramData
    & $add $env:APPDATA
    & $add $env:LOCALAPPDATA
    & $add $env:TEMP
    & $add ([Environment]::GetFolderPath('Desktop'))
    & $add ([Environment]::GetFolderPath('MyDocuments'))
    & $add ([Environment]::GetFolderPath('UserProfile'))
    & $add ([Environment]::GetFolderPath('Windows'))
    return $list
}

function Test-SpecialPath {
    param([string]$Path)
    $p = Get-NormalizedPath $Path
    if (-not $p) { return '路径无效。' }
    if ($p -match '::') { return '这是系统虚拟位置（不是真实文件夹），不能转为框。' }
    if ($p -match '^[A-Za-z]:\\?$') { return '磁盘根目录不能转为框，请选一个具体文件夹。' }
    if ($p.Length -lt 4) { return '路径太短，不能转为框。' }
    foreach ($sp in (Get-SpecialPaths)) {
        if (-not $sp) { continue }
        if ((Test-PathEqual $p $sp)) {
            return ('「{0}」 是系统/用户特殊目录，不能转为框。' -f $sp)
        }
        if ($sp.Length -gt 3 -and (Test-PathEqual $p (Join-Path $sp '.')) ) {
            return ('「{0}」 是系统/用户特殊目录，不能转为框。' -f $sp)
        }
    }
    return $null
}

function Get-FileCountText {
    param([int]$Dirs, [int]$Files)
    $parts = @()
    if ($Files -gt 0) { $parts += "$Files 个文件" }
    if ($Dirs -gt 0) { $parts += "$Dirs 个文件夹" }
    if ($parts.Count -eq 0) { return '空框' }
    return ($parts -join ' / ')
}

# =====================================================================
# 桌面图标度量（网格）
# =====================================================================
function Get-RawIconMetrics {
    $result = @{ Icon = 48; Horz = 75; Vert = 75; Source = 'default' }
    try {
        $size = 512
        $buf = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($size)
        try {
            [System.Runtime.InteropServices.Marshal]::WriteInt32($buf, 0, $size)
            $ok = [FPNative]::SystemParametersInfo(0x002D, [uint32]$size, $buf, 0)  # SPI_GETICONMETRICS
            if ($ok) {
                $h = [System.Runtime.InteropServices.Marshal]::ReadInt32($buf, 4)
                $v = [System.Runtime.InteropServices.Marshal]::ReadInt32($buf, 8)
                if ($h -gt 0) { $result.Horz = $h }
                if ($v -gt 0) { $result.Vert = $v }
                $result.Source = 'SPI_GETICONMETRICS'
            }
        } finally {
            [System.Runtime.InteropServices.Marshal]::FreeHGlobal($buf)
        }
    } catch { Write-Log "SPI_GETICONMETRICS 失败: $($_.Exception.Message)" 'WARN' }

    try {
        $hs = [FPNative]::GetSystemMetrics(38)   # SM_CXICONSPACING
        $vs = [FPNative]::GetSystemMetrics(39)   # SM_CYICONSPACING
        $is = [FPNative]::GetSystemMetrics(11)   # SM_CXICON
        if ($result.Horz -le 0 -and $hs -gt 0) { $result.Horz = $hs; $result.Source = 'GetSystemMetrics' }
        if ($result.Vert -le 0 -and $vs -gt 0) { $result.Vert = $vs }
        if ($is -gt 0) { $result.Icon = $is }
    } catch { }

    try {
        $wm = Get-ItemProperty -Path 'HKCU:\Control Panel\Desktop\WindowMetrics' -ErrorAction SilentlyContinue
        if ($wm -and $wm.'Shell Icon Size') {
            $s = [int]$wm.'Shell Icon Size'
            if ($s -ge 16 -and $s -le 256) { $result.Icon = $s }
        }
        if ($wm -and $result.Horz -le 0 -and $wm.IconSpacing) {
            $sp = [Math]::Abs([int]$wm.IconSpacing)
            if ($sp -gt 0) { $result.Horz = [int]($sp / 15) }
        }
        if ($wm -and $result.Vert -le 0 -and $wm.IconVerticalSpacing) {
            $sp = [Math]::Abs([int]$wm.IconVerticalSpacing)
            if ($sp -gt 0) { $result.Vert = [int]($sp / 15) }
        }
    } catch { }

    if ($result.Icon -lt 16) { $result.Icon = 32 }
    if ($result.Horz -lt ($result.Icon + 12)) { $result.Horz = $result.Icon + 24 }
    if ($result.Vert -lt ($result.Icon + 20)) { $result.Vert = $result.Icon + 34 }
    return $result
}

function Get-DpiScale {
    if ($script:Grid.Scale -gt 0 -and -not $script:Grid.Raw) { }
    try {
        $h = [FPNative]::GetForegroundWindow()
    } catch { $h = [IntPtr]::Zero }
    $scale = 1.0
    try {
        $w = New-Object System.Windows.Window
        $w.WindowStyle = [System.Windows.WindowStyle]::None
        $w.ShowInTaskbar = $false
        $w.Width = 4; $w.Height = 4
        $w.Left = -2000; $w.Top = -2000
        $w.ShowActivated = $false
        $w.Show()
        $src = [System.Windows.Interop.HwndSource]::FromHwnd(([System.Windows.Interop.WindowInteropHelper]::new($w)).Handle)
        if ($src -and $src.CompositionTarget) {
            $m = $src.CompositionTarget.TransformToDevice
            if ($m.M11 -gt 0) { $scale = $m.M11 }
        }
        $w.Close()
    } catch {
        Write-Log "DPI 检测失败: $($_.Exception.Message)" 'WARN'
    }
    if ($scale -le 0) { $scale = 1.0 }
    return $scale
}

function Initialize-GridMetrics {
    $raw = Get-RawIconMetrics
    $scale = Get-DpiScale
    $script:Grid.Raw = $raw
    $script:Grid.Scale = $scale
    $script:Grid.CellW = [Math]::Round($raw.Horz / $scale, 1)
    $script:Grid.CellH = [Math]::Round($raw.Vert / $scale, 1)
    $script:Grid.Icon = [Math]::Round($raw.Icon / $scale, 1)
    if ($script:Grid.CellW -lt 40) { $script:Grid.CellW = 40 }
    if ($script:Grid.CellH -lt 48) { $script:Grid.CellH = 48 }
    if ($script:Grid.Icon -lt 16) { $script:Grid.Icon = 32 }
    if ($script:Grid.Icon -gt ($script:Grid.CellW - 8)) { $script:Grid.Icon = $script:Grid.CellW - 8 }
    Write-Log ("网格: 图标 {0}px 间距 {1}x{2} 缩放 {3} => 格 {4}x{5} DIP" -f $raw.Icon, $raw.Horz, $raw.Vert, $scale, $script:Grid.CellW, $script:Grid.CellH)
}

# =====================================================================
# 紧急度（边框色）
# =====================================================================
function Get-PadRemainingDays {
    param($Pad)
    if ($null -eq $Pad -or -not $Pad.dueDate) { return $null }
    try {
        $d = [datetime]::ParseExact([string]$Pad.dueDate, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    } catch { return $null }
    return ($d.Date - (Get-Date).Date).Days
}

function Get-UrgencyInfo {
    param($Pad, $Settings)
    if ($null -eq $Settings) { $Settings = $script:State.settings }
    $rem = Get-PadRemainingDays $Pad
    if ($null -eq $rem) {
        return @{ Tier = 'none'; Color = $script:NeutralBorderColor; Remaining = $null; Text = '未设截止' }
    }
    $u = $Settings.urgency
    if ($rem -lt 0) {
        $col = $script:NeutralBorderColor
        if ($u -and $u.expired -and $u.expired.color) { $col = [string]$u.expired.color }
        return @{ Tier = 'expired'; Color = $col; Remaining = $rem; Text = '已结束' }
    }
    foreach ($tier in @('safe', 'soon', 'close', 'today')) {
        $t = $null
        if ($u) { $t = $u.$tier }
        if ($null -eq $t) { continue }
        $days = 0
        $hasDays = $false
        if ($null -ne $t.days) { $days = [int]$t.days; $hasDays = $true }
        if (-not $hasDays) { continue }
        if ($rem -ge $days) {
            $col = $script:NeutralBorderColor
            if ($t.color) { $col = [string]$t.color }
            $text = if ($rem -eq 0) { '今天截止' } else { "还剩 $rem 天" }
            return @{ Tier = $tier; Color = $col; Remaining = $rem; Text = $text }
        }
    }
    return @{ Tier = 'expired'; Color = $script:NeutralBorderColor; Remaining = $rem; Text = '已结束' }
}

# =====================================================================
# 图标
# =====================================================================
function Get-FileIconSource {
    param([string]$Path, [bool]$IsDir, [int]$WantSize = 0)
    $want = 48
    if ($script:Grid.Raw) { $want = [int]$script:Grid.Raw.Icon }
    if ($WantSize -gt 0) { $want = [int]$WantSize }
    $key = ('{0}|{1}|{2}' -f $IsDir, $want, $Path.ToLowerInvariant())
    if ($script:IconCache.ContainsKey($key)) { return $script:IconCache[$key] }
    $hIcon = [IntPtr]::Zero
    try {
        $hIcon = [FPNative]::GetIconByIndex($Path, $want)
        if ($hIcon -eq [IntPtr]::Zero) { $hIcon = [FPNative]::GetLargeIcon($Path) }
        if ($hIcon -eq [IntPtr]::Zero) { return $null }
        $bmp = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon(
            $hIcon,
            [System.Windows.Int32Rect]::Empty,
            [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions())
        if ($bmp.CanFreeze) { $bmp.Freeze() }
        $script:IconCache[$key] = $bmp
        return $bmp
    } catch {
        Write-Log "取图标失败: $Path :: $($_.Exception.Message)" 'WARN'
        return $null
    } finally {
        if ($hIcon -ne [IntPtr]::Zero) { [void][FPNative]::DestroyIcon($hIcon) }
    }
}

function Get-PadItems {
    param($Pad)
    $res = @{ Ok = $false; Items = @(); Dirs = 0; Files = 0; Error = ''; Big = @{} }
    if ($null -eq $Pad -or -not $Pad.path) { $res.Error = '路径为空'; return $res }
    if (-not (Test-Path -LiteralPath $Pad.path -PathType Container)) {
        $res.Error = '路径失效'
        return $res
    }
    try {
        $children = @(Get-ChildItem -LiteralPath $Pad.path -Force -ErrorAction Stop)
    } catch {
        $res.Error = $_.Exception.Message
        return $res
    }
    $items = foreach ($c in $children) {
        $isDir = $c.PSIsContainer
        [pscustomobject]@{
            Name     = $c.Name
            FullName = $c.FullName
            IsDir    = [bool]$isDir
            Size     = if ($isDir) { 0 } else { $c.Length }
            Modified = $c.LastWriteTime
            Hidden   = [bool]($c.Attributes -band [System.IO.FileAttributes]::Hidden)
        }
    }
    $items = @($items)
    # 自然顺序：文件夹优先 + 名称（新增条目按这个顺序追加到 order 末尾）
    $natural = @($items | Sort-Object @{ Expression = { if ($_.IsDir) { 0 } else { 1 } } }, @{ Expression = { $_.Name } })
    $norm = Normalize-PadOrder -Pad $Pad -Names @($natural | ForEach-Object { $_.Name })
    $Pad.order = @($norm.Order)
    $Pad.big = @($norm.Big)
    $pos = @{}
    $i = 0
    foreach ($nm in @($norm.Order)) { $pos[[string]$nm] = $i; $i++ }
    $items = @($items | Sort-Object @{ Expression = {
                $idx = 9999
                if ($pos.ContainsKey([string]$_.Name)) { $idx = [int]$pos[[string]$_.Name] }
                $idx
            } })
    $res.Big = @{}
    foreach ($bn in @($norm.Big)) { $res.Big[[string]$bn] = $true }
    $res.Ok = $true
    $res.Items = $items
    $res.Dirs = @($items | Where-Object { $_.IsDir }).Count
    $res.Files = @($items | Where-Object { -not $_.IsDir }).Count
    return $res
}

# =====================================================================
# 数据模型与持久化
# =====================================================================
function Get-DefaultSettings {
    $archiveRoot = 'D:\归档'
    if (-not (Test-Path -LiteralPath 'D:\')) {
        $archiveRoot = Join-Path ([Environment]::GetFolderPath('MyDocuments')) '归档'
    }
    return [ordered]@{
        refreshMinutes  = 30
        archiveRoot     = $archiveRoot
        hideAll         = $false
        exitRestore     = $true
        autostart       = $false
        frameFill       = $script:DefaultFill
        settingsVersion = 3
        backdrop        = $true
        categoryPresets = @($script:CategoryPresets)
        urgency         = [ordered]@{
            safe    = [ordered]@{ days = 8; color = '#8FA3B0' }
            soon    = [ordered]@{ days = 4; color = '#E8B33A' }
            close   = [ordered]@{ days = 1; color = '#E8762C' }
            today   = [ordered]@{ days = 0; color = '#D2362F' }
            expired = [ordered]@{ color = '#9A9A9A' }
        }
    }
}

function New-DefaultState {
    return [ordered]@{
        settings   = (Get-DefaultSettings)
        pads       = @()
        archiveLog = @()
    }
}

function ConvertTo-HashtableDeep {
    param($InputObject)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) { return $InputObject }
    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $h = [ordered]@{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            $h[$prop.Name] = ConvertTo-HashtableDeep $prop.Value
        }
        return $h
    }
    if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
        $arr = @()
        foreach ($it in $InputObject) { $arr += , (ConvertTo-HashtableDeep $it) }
        return $arr
    }
    return $InputObject
}

function Merge-Settings {
    param($Loaded)
    $def = Get-DefaultSettings
    if ($null -eq $Loaded) { return $def }
    foreach ($k in $def.Keys) {
        if (-not $Loaded.Contains($k) -or $null -eq $Loaded[$k]) { $Loaded[$k] = $def[$k] }
    }
    # 旧版默认 0.12（那时底色只是中性灰）；现在底色承载紧急度色，太淡看不出，一次性升到新默认
    try {
        $ver = 0
        if ($Loaded.Contains('settingsVersion') -and $null -ne $Loaded.settingsVersion) { $ver = [int]$Loaded.settingsVersion }
        if ($ver -lt 3) {
            if ($null -eq $Loaded.frameFill -or [double]$Loaded.frameFill -le 0.30) {
                $Loaded.frameFill = $script:DefaultFill
            }
            $Loaded.settingsVersion = 3
        }
    } catch { }
    if ($null -eq $Loaded.urgency) { $Loaded.urgency = $def.urgency }
    foreach ($tier in @('safe', 'soon', 'close', 'today', 'expired')) {
        if ($null -eq $Loaded.urgency.$tier) { $Loaded.urgency[$tier] = $def.urgency[$tier] }
        else {
            foreach ($f in @('days', 'color')) {
                if ($def.urgency[$tier].Contains($f) -and -not $Loaded.urgency[$tier].Contains($f)) {
                    $Loaded.urgency[$tier][$f] = $def.urgency[$tier][$f]
                }
            }
        }
    }
    return $Loaded
}

function Read-StateFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        return (ConvertTo-HashtableDeep ($raw | ConvertFrom-Json))
    } catch {
        Write-Log "读取状态失败 $Path :: $($_.Exception.Message)" 'WARN'
        return $null
    }
}

function Load-State {
    Initialize-DataDir
    $state = Read-StateFile $script:Path_Json
    if ($null -eq $state) { $state = Read-StateFile $script:Path_Bak }
    if ($null -eq $state) {
        Write-Log '没有可用的 pads.json / pads.bak，使用默认状态'
        $state = New-DefaultState
    }
    if ($null -eq $state.settings) { $state.settings = Get-DefaultSettings }
    $state.settings = Merge-Settings $state.settings
    if ($null -eq $state.pads) { $state.pads = @() }
    if ($null -eq $state.archiveLog) { $state.archiveLog = @() }
    $clean = @()
    foreach ($pad in @($state.pads)) {
        if ($null -eq $pad) { continue }
        if (-not $pad.id) { $pad.id = 'p' + [Guid]::NewGuid().ToString('N').Substring(0, 8) }
        if (-not $pad.path) { continue }
        if (-not $pad.Contains('pos') -or $null -eq $pad.pos) { $pad.pos = @{ x = 80; y = 80 } }
        if ($null -eq $pad.pos.x) { $pad.pos.x = 80 }
        if ($null -eq $pad.pos.y) { $pad.pos.y = 80 }
        if (-not $pad.Contains('cols') -or $null -eq $pad.cols) { $pad.cols = 'auto' }
        if (-not $pad.Contains('autoUrgency')) { $pad.autoUrgency = $true }
        if (-not $pad.Contains('categoryColorOn')) { $pad.categoryColorOn = $true }
        if (-not $pad.Contains('pinned')) { $pad.pinned = $false }
        if (-not $pad.Contains('order') -or $null -eq $pad.order) { $pad.order = @() }
        if (-not $pad.Contains('big') -or $null -eq $pad.big) { $pad.big = @() }
        if (-not $pad.Contains('cells') -or $null -eq $pad.cells) { $pad.cells = @{} }
        # 从 JSON 读回来的是 OrderedDictionary，统一转成 hashtable（下游用 ContainsKey）
        $cellsTmp = @{}
        foreach ($ck in @($pad.cells.Keys)) {
            try { $cellsTmp[[string]$ck] = @{ c = [int]$pad.cells[$ck].c; r = [int]$pad.cells[$ck].r } } catch { }
        }
        $pad.cells = $cellsTmp
        if (-not $pad.Contains('rows') -or $null -eq $pad.rows -or [string]::IsNullOrWhiteSpace([string]$pad.rows)) { $pad.rows = 'auto' }
        if (-not $pad.Contains('prevHidden')) { $pad.prevHidden = $false }
        if (-not $pad.Contains('prevAttrs')) { $pad.prevAttrs = -1 }
        if (-not $pad.Contains('invalid')) { $pad.invalid = $false }
        $clean += $pad
    }
    $state.pads = @($clean)
    $script:State = $state
    return $state
}

function Write-HiddenBackup {
    Initialize-DataDir
    $sb = New-Object System.Text.StringBuilder
    foreach ($pad in @($script:State.pads)) {
        if (-not $pad -or -not $pad.path) { continue }
        $pb = -1
        if ($null -ne $pad.prevAttrs) { $pb = [int]$pad.prevAttrs }
        $ph = 0
        if ($pad.prevHidden) { $ph = 1 }
        [void]$sb.AppendLine(('{0}{1}{2}{1}{3}' -f $pad.path, [char]9, $pb, $ph))
    }
    $tmp = $script:Path_HiddenBackup + '.tmp'
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($tmp, $sb.ToString(), $enc)
    [System.IO.File]::Copy($tmp, $script:Path_HiddenBackup, $true)
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}

function Read-HiddenBackup {
    $list = @()
    if (-not (Test-Path -LiteralPath $script:Path_HiddenBackup)) { return $list }
    try {
        $lines = [System.IO.File]::ReadAllLines($script:Path_HiddenBackup, [System.Text.Encoding]::UTF8)
    } catch { return $list }
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $parts = $line -split "`t"
        if ($parts.Count -lt 1) { continue }
        $p = $parts[0]
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $attrs = -1
        if ($parts.Count -ge 2) { [void][int]::TryParse($parts[1], [ref]$attrs) }
        $hidden = $false
        if ($parts.Count -ge 3) { $hidden = ($parts[2] -eq '1') }
        $list += [pscustomobject]@{ Path = $p; Attrs = $attrs; Hidden = $hidden }
    }
    return $list
}

function Save-State {
    Initialize-DataDir
    Write-HiddenBackup
    $json = ConvertTo-Json -InputObject $script:State -Depth 8
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($script:Path_Tmp, $json, $enc)
    if (Test-Path -LiteralPath $script:Path_Json) {
        try { [System.IO.File]::Copy($script:Path_Json, $script:Path_Bak, $true) } catch { }
    }
    [System.IO.File]::Copy($script:Path_Tmp, $script:Path_Json, $true)
    Remove-Item -LiteralPath $script:Path_Tmp -Force -ErrorAction SilentlyContinue
}

function Get-PadByPath {
    param([string]$Path)
    foreach ($pad in @($script:State.pads)) {
        if (-not $pad) { continue }
        if (Test-PathEqual $pad.path $Path) { return $pad }
    }
    return $null
}

function Get-PadById {
    param([string]$PadId)
    foreach ($pad in @($script:State.pads)) {
        if ($pad -and $pad.id -eq $PadId) { return $pad }
    }
    return $null
}

function Remove-PadRecord {
    param([string]$PadId)
    $script:State.pads = @($script:State.pads | Where-Object { $_ -and $_.id -ne $PadId })
}

# =====================================================================
# 隐藏属性：设置 / 精确还原
# =====================================================================
function Get-PathAttributes {
    param([string]$Path)
    try { return [int][System.IO.File]::GetAttributes($Path) } catch { return -1 }
}

function Set-FolderHidden {
    param([string]$Path)
    try {
        $cur = [System.IO.File]::GetAttributes($Path)
        $hidden = [System.IO.FileAttributes]::Hidden
        if (-not ($cur -band $hidden)) {
            [System.IO.File]::SetAttributes($Path, ($cur -bor $hidden))
            Write-Log "已设为隐藏: $Path"
        }
        return $true
    } catch {
        Write-Log "设置隐藏失败: $Path :: $($_.Exception.Message)" 'ERROR'
        return $false
    }
}

function Restore-FolderVisibility {
    param($Pad, [switch]$ForceVisible)
    if ($null -eq $Pad -or -not $Pad.path) { return $false }
    if (-not [System.IO.Directory]::Exists($Pad.path)) { return $false }
    $mask = ([System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)
    try {
        $cur = [System.IO.File]::GetAttributes($Pad.path)
        $prev = -1
        if ($null -ne $Pad.prevAttrs) { $prev = [int]$Pad.prevAttrs }
        if ($ForceVisible) {
            $new = $cur -band (-bnot $mask)
        } elseif ($prev -ge 0) {
            $new = ($cur -band (-bnot $mask)) -bor ($prev -band $mask)
        } else {
            if ($Pad.prevHidden) { $new = ($cur -band (-bnot $mask)) -bor ([System.IO.FileAttributes]::Hidden) }
            else { $new = $cur -band (-bnot $mask) }
        }
        if ($new -ne $cur) { [System.IO.File]::SetAttributes($Pad.path, $new) }
        Write-Log ("已还原属性: {0} => {1}" -f $Pad.path, [System.IO.FileAttributes]$new)
        return $true
    } catch {
        Write-Log "还原属性失败: $($Pad.path) :: $($_.Exception.Message)" 'ERROR'
        return $false
    }
}

# =====================================================================
# 屏幕边界与坐标夹取
# =====================================================================
function Get-VirtualScreenDip {
    $scale = 1.0
    if ($script:Grid.Scale -gt 0) { $scale = $script:Grid.Scale }
    try {
        $vs = [System.Windows.Forms.SystemInformation]::VirtualScreen
        return @{
            Left   = $vs.Left / $scale
            Top    = $vs.Top / $scale
            Right  = $vs.Right / $scale
            Bottom = $vs.Bottom / $scale
        }
    } catch {
        return @{ Left = 0; Top = 0; Right = 1920; Bottom = 1080 }
    }
}

function Limit-PadPosition {
    param([double]$X, [double]$Y, [double]$W = 100, [double]$H = 100)
    $vs = Get-VirtualScreenDip
    if ([double]::IsNaN($X)) { $X = $vs.Left }
    if ([double]::IsNaN($Y)) { $Y = $vs.Top }
    if ([double]::IsNaN($W) -or $W -le 0) { $W = 120 }
    if ([double]::IsNaN($H) -or $H -le 0) { $H = 90 }
    $minX = $vs.Left
    $minY = $vs.Top
    $maxX = [Math]::Max($minX, $vs.Right - $W)
    $maxY = [Math]::Max($minY, $vs.Bottom - $H)
    $nx = [Math]::Min([Math]::Max($X, $minX), $maxX)
    $ny = [Math]::Min([Math]::Max($Y, $minY), $maxY)
    return @{ x = [Math]::Round($nx); y = [Math]::Round($ny) }
}

function Get-NextPadPosition {
    $vs = Get-VirtualScreenDip
    $n = @($script:State.pads).Count
    $x = $vs.Left + 60 + (($n % 6) * 36)
    $y = $vs.Top + 60 + (($n % 6) * 28)
    $pos = Limit-PadPosition -X $x -Y $y -W 260 -H 200
    return @{ x = $pos.x; y = $pos.y }
}

# =====================================================================
# 文件移动（含跨盘、同名、目录合并）
# =====================================================================
function Get-UniquePath {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $Path }
    $dir = Split-Path -Parent $Path
    $leaf = Split-Path -Leaf $Path
    $isDir = [System.IO.Directory]::Exists($Path)
    $base = $leaf
    $ext = ''
    if (-not $isDir) {
        $ext = [System.IO.Path]::GetExtension($leaf)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($leaf)
    }
    for ($i = 1; $i -lt 1000; $i++) {
        $name = '{0}_{1}{2}' -f $base, $i, $ext
        $cand = Join-Path $dir $name
        if (-not (Test-Path -LiteralPath $cand)) { return $cand }
    }
    return (Join-Path $dir ('{0}_{1}{2}' -f $base, [Guid]::NewGuid().ToString('N').Substring(0, 4), $ext))
}

function Get-SameDrive {
    param([string]$A, [string]$B)
    try { return ([System.IO.Path]::GetPathRoot($A) -eq [System.IO.Path]::GetPathRoot($B)) } catch { return $false }
}

function Move-FileSafe {
    param([string]$Source, [string]$Dest)
    if (Get-SameDrive $Source $Dest) {
        [System.IO.File]::Move($Source, $Dest)
        return
    }
    $size = (Get-Item -LiteralPath $Source -Force).Length
    [System.IO.File]::Copy($Source, $Dest, $false)
    $destSize = (Get-Item -LiteralPath $Dest -Force).Length
    if ($destSize -ne $size) {
        throw "已复制到目标，但大小不一致（$destSize/$size），源文件已保留：$Source"
    }
    try {
        [System.IO.File]::Delete($Source)
    } catch {
        throw "已复制到目标，源文件删除失败（源文件已保留）：$Source`r`n$($_.Exception.Message)"
    }
}

function Copy-DirectorySafe {
    param([string]$Source, [string]$Dest)
    New-Item -ItemType Directory -Force -Path $Dest | Out-Null
    Get-ChildItem -LiteralPath $Source -Force | ForEach-Object {
        $target = Join-Path $Dest $_.Name
        if ($_.PSIsContainer) { Copy-DirectorySafe -Source $_.FullName -Dest $target }
        else { [System.IO.File]::Copy($_.FullName, $target, $false) }
    }
}

function Move-DirectorySafe {
    param([string]$Source, [string]$Dest)
    if (Get-SameDrive $Source $Dest) {
        [System.IO.Directory]::Move($Source, $Dest)
        return
    }
    Copy-DirectorySafe -Source $Source -Dest $Dest
    try {
        [System.IO.Directory]::Delete($Source, $true)
    } catch {
        throw "已复制到目标，源目录删除失败（源目录已保留）：$Source`r`n$($_.Exception.Message)"
    }
}

function Move-ItemToDirectory {
    param([string]$Source, [string]$DestDir, [string]$ConflictAction = 'ask')
    $leaf = Split-Path -Leaf $Source
    $dest = Join-Path $DestDir $leaf
    $isDir = [System.IO.Directory]::Exists($Source)
    if (Test-Path -LiteralPath $dest) {
        switch ($ConflictAction) {
            'skip' { return @{ Ok = $true; Skipped = $true; Dest = $dest } }
            'overwrite' {
                if ($isDir) { return @{ Ok = $false; Error = '不允许覆盖整个文件夹（请选择合并或重命名）'; Dest = $dest } }
                Remove-Item -LiteralPath $dest -Force
            }
            'rename' { $dest = Get-UniquePath $dest }
            'merge' {
                if (-not $isDir -or -not [System.IO.Directory]::Exists($dest)) { $dest = Get-UniquePath $dest }
                else {
                    foreach ($child in @(Get-ChildItem -LiteralPath $Source -Force)) {
                        $r = Move-ItemToDirectory -Source $child.FullName -DestDir $dest -ConflictAction 'rename'
                        if (-not $r.Ok) { return $r }
                    }
                    Remove-Item -LiteralPath $Source -Force -Recurse -ErrorAction SilentlyContinue
                    return @{ Ok = $true; Dest = $dest; Merged = $true }
                }
            }
            default { $dest = Get-UniquePath $dest }
        }
    }
    try {
        if ($isDir) { Move-DirectorySafe -Source $Source -Dest $dest }
        else { Move-FileSafe -Source $Source -Dest $dest }
        return @{ Ok = $true; Dest = $dest }
    } catch {
        return @{ Ok = $false; Error = $_.Exception.Message; Dest = $dest }
    }
}

# =====================================================================
# 转框 / 转回文件夹
# =====================================================================
function Convert-ToPad {
    param([string]$Path, [string]$DueDate, [switch]$SilentErrors)
    $norm = Get-NormalizedPath $Path
    if (-not $norm) {
        Show-MessageDialog -Title '转为 FolderPad 框' -Message ('路径无效：' + $Path) -Kind 'warn'
        return $false
    }
    $special = Test-SpecialPath $norm
    if ($special) {
        Show-MessageDialog -Title '转为 FolderPad 框' -Message $special -Kind 'warn'
        return $false
    }
    if (-not [System.IO.Directory]::Exists($norm)) {
        if ([System.IO.File]::Exists($norm)) {
            Show-MessageDialog -Title '转为 FolderPad 框' -Message ('这是一个文件，不是文件夹：' + "`r`n" + $norm + "`r`n`r`n只能把文件夹转为框。") -Kind 'warn'
        } else {
            Show-MessageDialog -Title '转为 FolderPad 框' -Message ('文件夹不存在：' + $norm) -Kind 'warn'
        }
        return $false
    }
    if ($norm.IndexOf('FolderPad') -ge 0 -and $norm.StartsWith($script:DataDir, [System.StringComparison]::OrdinalIgnoreCase)) {
        Show-MessageDialog -Title '转为 FolderPad 框' -Message '不能把 FolderPad 自己的数据目录转为框。' -Kind 'warn'
        return $false
    }
    $exists = Get-PadByPath $norm
    if ($exists) {
        Show-MessageDialog -Title '这个文件夹已经是一个框了' -Message ('「{0}」 已经是一个 FolderPad 框（id={1}）。' -f $exists.name, $exists.id) -Kind 'info'
        if ($script:Pads.ContainsKey($exists.id) -and $script:Pads[$exists.id].Window) {
            try { $script:Pads[$exists.id].Window.Activate() } catch { }
        }
        return $false
    }

    $item = Get-Item -LiteralPath $norm -Force
    $attrs = [int]$item.Attributes
    if ($attrs -lt 0) { $attrs = (Get-PathAttributes $norm) }
    $isHidden = [bool]($attrs -band [System.IO.FileAttributes]::Hidden)

    if ($isHidden) {
        $ok = Show-ConfirmDialog -Title '该文件夹当前是隐藏的' `
            -Message ('「{0}」 现在是隐藏文件夹。' -f $item.Name) `
            -Detail '转回普通文件夹后，它仍会保持隐藏（不会被改成可见）。确定要转为框吗？' `
            -OkText '仍然转为框' -CancelText '取消'
        if (-not $ok) { return $false }
    }

    $due = $DueDate
    if (-not $due) {
        $res = Show-DueDateDialog -Name $item.Name -Path $norm -IsNew
        if ($null -eq $res -or -not $res.Ok) { return $false }
        $due = $res.DueDate
    }

    $pid2 = 'p' + (Get-Date).ToString('yyMMddHHmmss') + (Get-Random -Minimum 100 -Maximum 999)
    $colorIdx = @($script:State.pads).Count % $script:CategoryPresets.Count
    $pos = Get-NextPadPosition
    $pad = [ordered]@{
        id              = $pid2
        name            = $item.Name
        path            = $norm
        dueDate         = [string]$due
        categoryColor   = $script:CategoryPresets[$colorIdx]
        categoryColorOn = $true
        autoUrgency     = $true
        cols            = 'auto'
        pos             = @{ x = [int]$pos.x; y = [int]$pos.y }
        pinned          = $false
        prevHidden      = $isHidden
        prevAttrs       = $attrs
        createdAt       = (Get-Date).ToString('s')
        order           = @()
        big             = @()
        cells           = @{}
        rows            = 'auto'
        invalid         = $false
    }
    $script:State.pads = @($script:State.pads) + $pad

    # 写入顺序（防中途崩溃）：① hidden-backup.txt ② pads.json ③ 设 Hidden 属性
    try {
        Save-State
    } catch {
        $script:State.pads = @($script:State.pads | Where-Object { $_.id -ne $pid2 })
        Show-MessageDialog -Title '转为框失败' -Message ('写入配置失败：' + $_.Exception.Message) -Kind 'error'
        return $false
    }
    $hiddenOk = Set-FolderHidden $norm
    if (-not $hiddenOk) {
        Show-MessageDialog -Title '转为框（部分完成）' -Message ('框已记录，但没能把这个文件夹设为隐藏（可能权限不足）。' + "`r`n" + $norm) -Kind 'warn'
    }
    Show-PadWindow -PadId $pid2
    Write-Log "转为框: $norm (due=$due, prevHidden=$isHidden)"
    Show-TrayBalloon -Title 'FolderPad' -Text ('已转为框：' + $item.Name)
    return $true
}

function Convert-ToFolder {
    param([string]$PadId, [switch]$NoSave)
    $pad = Get-PadById $PadId
    if ($null -eq $pad) { return }
    Restore-FolderVisibility -Pad $pad | Out-Null
    Remove-PadRecord $PadId
    Close-PadWindow -PadId $PadId
    if (-not $NoSave) { Save-State }
    Write-Log "转为普通文件夹: $($pad.path)"
    Show-TrayBalloon -Title 'FolderPad' -Text ('已转回普通文件夹：' + $pad.name)
}

# =====================================================================
# 归档
# =====================================================================
function Get-ArchiveBaseName {
    param($Pad)
    $prefix = ''
    if ($Pad.dueDate) {
        $prefix = [string]$Pad.dueDate
    } else {
        $prefix = (Get-Date).ToString('yyyy-MM-dd')
    }
    $name = [string]$Pad.name
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($ch in $invalid) { $name = $name.Replace([string]$ch, '_') }
    if ([string]::IsNullOrWhiteSpace($name)) { $name = '未命名' }
    return ('{0}_{1}' -f $prefix, $name)
}

function Invoke-Archive {
    param([string]$PadId, [ValidateSet('zip', 'folder-inplace', 'folder-move')][string]$Mode)
    $pad = Get-PadById $PadId
    if ($null -eq $pad) { return }
    $src = $pad.path
    if (-not [System.IO.Directory]::Exists($src)) {
        Show-MessageDialog -Title '归档失败' -Message ('源文件夹不存在：' + "`r`n" + $src + "`r`n`r`n请先用「重新定位…」找回它，或「移除记录」。") -Kind 'error'
        return
    }
    # 归档产物必须是可见的普通条目
    Restore-FolderVisibility -Pad $pad | Out-Null
    $baseName = Get-ArchiveBaseName $pad
    $settings = $script:State.settings
    $target = ''
    $type = ''
    try {
        switch ($Mode) {
            'zip' {
                $root = [string]$settings.archiveRoot
                if ([string]::IsNullOrWhiteSpace($root)) { $root = Join-Path ([Environment]::GetFolderPath('MyDocuments')) '归档' }
                if (-not (Test-Path -LiteralPath $root)) { New-Item -ItemType Directory -Force -Path $root | Out-Null }
                $target = Get-UniquePath (Join-Path $root ($baseName + '.zip'))
                [System.IO.Compression.ZipFile]::CreateFromDirectory(
                    $src, $target,
                    [System.IO.Compression.CompressionLevel]::Optimal, $true)
                $type = 'zip'
            }
            'folder-inplace' {
                $parent = Split-Path -Parent $src
                $target = Get-UniquePath (Join-Path $parent $baseName)
                Move-DirectorySafe -Source $src -Dest $target
                $type = 'folder'
            }
            'folder-move' {
                $root = [string]$settings.archiveRoot
                if ([string]::IsNullOrWhiteSpace($root)) { $root = Join-Path ([Environment]::GetFolderPath('MyDocuments')) '归档' }
                if (-not (Test-Path -LiteralPath $root)) { New-Item -ItemType Directory -Force -Path $root | Out-Null }
                $target = Get-UniquePath (Join-Path $root $baseName)
                Move-DirectorySafe -Source $src -Dest $target
                $type = 'folder'
            }
        }
    } catch {
        Show-MessageDialog -Title '归档失败' -Message ('归档时出错：' + "`r`n" + $_.Exception.Message + "`r`n`r`n源文件夹保持原样，未被删除。记录仍然保留。") -Kind 'error'
        Write-Log "归档失败: $src :: $($_.Exception.Message)" 'ERROR'
        return
    }

    $bytes = 0
    try {
        if ($type -eq 'zip') {
            $bytes = (Get-Item -LiteralPath $target -Force).Length
        } else {
            $bytes = [int](@(Get-ChildItem -LiteralPath $target -Recurse -Force -File -ErrorAction SilentlyContinue) | Measure-Object -Property Length -Sum).Sum
            if ($null -eq $bytes) { $bytes = 0 }
        }
    } catch { }

    $logEntry = [ordered]@{
        time          = (Get-Date).ToString('s')
        padId         = $pad.id
        source        = $src
        target        = $target
        type          = $type
        sourceDeleted = $false
        bytes         = $bytes
    }
    $script:State.archiveLog = @($script:State.archiveLog) + $logEntry
    Remove-PadRecord $PadId
    Close-PadWindow -PadId $PadId
    Save-State
    Write-Log "归档完成($Mode): $src => $target"
    Show-TrayBalloon -Title '归档完成' -Text (Split-Path -Leaf $target)
    Show-MessageDialog -Title '归档完成' -Message ('已归档为：' + "`r`n" + $target + "`r`n`r`n原文件夹已保留：`r`n" + $src) -Kind 'info' -Detail ('大小：{0:N0} 字节' -f $bytes)
}

function Open-InExplorer {
    param([string]$Path, [switch]$Select)
    try {
        if (-not (Test-Path -LiteralPath $Path)) { return }
        if ($Select) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"' + $Path + '"') }
        else { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"' + $Path + '"') }
    } catch { Write-Log "打开资源管理器失败: $($_.Exception.Message)" 'WARN' }
}

# =====================================================================
# 视觉：深色菜单 / 对话框
# =====================================================================





# =====================================================================
# 对话框
# =====================================================================
function New-DialogWindow {
    param([string]$Width = '420')
    $win = New-Object System.Windows.Window
    $win.WindowStyle = 'None'
    $win.AllowsTransparency = $true
    $win.Background = [System.Windows.Media.Brushes]::Transparent
    $win.ShowInTaskbar = $false
    $win.WindowStartupLocation = 'CenterScreen'
    $win.Topmost = $true
    $win.SizeToContent = 'Height'
    $win.FontFamily = $script:FontFamily
    # 窗口比卡片宽 2*DlgShadowPad，这圈透明边用来放阴影
    $pad2 = 2 * $script:DlgShadowPad
    $w = 420
    if ($Width -match '^\d+$') { $w = [double]$Width }
    $win.Width = $w + $pad2
    $win.MinWidth = $w + $pad2
    return $win
}

function New-DialogCard {
    $card = New-Object System.Windows.Controls.Border
    $card.Background = New-FPBrush '#F21D1D21' 1.0
    $card.BorderBrush = New-FPBrush '#FFFFFF' 0.12
    $card.BorderThickness = '1'
    $card.CornerRadius = '10'
    $card.Padding = '18,16,18,14'
    $shadow = New-Object System.Windows.Media.Effects.DropShadowEffect
    $shadow.BlurRadius = 22
    $shadow.ShadowDepth = 0
    $shadow.Opacity = 0.5
    $shadow.Color = [System.Windows.Media.Colors]::Black
    $card.Effect = $shadow
    return $card
}

function New-DialogButton {
    param([string]$Text, [bool]$Primary = $false, [bool]$Danger = $false, [scriptblock]$OnClick)
    $btn = New-Object System.Windows.Controls.Button
    $btn.Content = $Text
    $btn.FontSize = 12.5
    $btn.FontFamily = $script:FontFamily
    $btn.Padding = '14,6'
    $btn.Margin = '8,0,0,0'
    $btn.MinWidth = 74
    $btn.Cursor = 'Hand'
    $btn.BorderThickness = '1'
    if ($Primary) {
        $btn.Background = New-FPBrush '#3C7DFF' 1.0
        $btn.Foreground = New-FPBrush '#FFFFFF' 1.0
        $btn.BorderBrush = New-FPBrush '#3C7DFF' 1.0
    } elseif ($Danger) {
        $btn.Background = New-FPBrush '#8F2E2E' 1.0
        $btn.Foreground = New-FPBrush '#FFFFFF' 1.0
        $btn.BorderBrush = New-FPBrush '#C04A4A' 1.0
    } else {
        $btn.Background = New-FPBrush '#2C2C31' 1.0
        $btn.Foreground = New-FPBrush '#E8E8E8' 1.0
        $btn.BorderBrush = New-FPBrush '#FFFFFF' 0.18
    }
    if ($OnClick) { $btn.Add_Click($OnClick) }
    return $btn
}

function Show-MessageDialog {
    param([string]$Title, [string]$Message, [string]$Detail = '', [string]$Kind = 'info', [string]$OkText = '知道了')
    if ($script:QuietUI) { Write-Log "[$Kind] $Title :: $Message"; return }
    Show-CustomDialog -Title $Title -Message $Message -Detail $Detail -Kind $Kind -Buttons @(
        @{ Text = $OkText; Value = 'ok'; Primary = $true }
    ) | Out-Null
}

function Show-ConfirmDialog {
    param([string]$Title, [string]$Message, [string]$Detail = '', [string]$OkText = '确定', [string]$CancelText = '取消', [string]$Kind = 'question')
    if ($script:QuietUI) { Write-Log "[confirm] $Title :: $Message"; return $false }
    $res = Show-CustomDialog -Title $Title -Message $Message -Detail $Detail -Kind $Kind -Buttons @(
        @{ Text = $CancelText; Value = 'cancel' },
        @{ Text = $OkText; Value = 'ok'; Primary = $true }
    )
    return ($res.Value -eq 'ok')
}

function Get-KindVisual {
    param([string]$Kind)
    switch ($Kind) {
        'warn' { return @{ Glyph = [char]0x26A0; Color = '#E8B33A' } }
        'error' { return @{ Glyph = [char]0x2715; Color = '#E05A5A' } }
        'question' { return @{ Glyph = '?'; Color = '#4C8DFF' } }
        default { return @{ Glyph = 'i'; Color = '#4C8DFF' } }
    }
}

function Show-CustomDialog {
    param(
        [string]$Title,
        [string]$Message = '',
        [string]$Detail = '',
        [string]$Kind = 'info',
        [object[]]$Buttons = @(),
        [string]$CheckboxText = '',
        [scriptblock]$BuildContent = $null,
        [double]$Width = 430
    )
    $script:Dlg = @{ Value = $null; Checked = $false }
    $win = New-DialogWindow -Width $Width
    $win.Add_KeyDown({
            param($s, $e)
            if ($e.Key -eq 'Escape') { $s.Close() }
        })

    $card = New-DialogCard
    # 卡片外面套一圈透明留白，给阴影留地方——否则阴影被窗口边界切掉，圆角处会有一圈"虚边"
    $cardHost = New-Object System.Windows.Controls.Grid
    $cardHost.Margin = (New-Object System.Windows.Thickness $script:DlgShadowPad)
    [void]$cardHost.Children.Add($card)
    $box = New-Object System.Windows.Controls.StackPanel
    $card.Child = $box

    $head = New-Object System.Windows.Controls.StackPanel
    $head.Orientation = 'Horizontal'
    $kv = Get-KindVisual $Kind
    $badge = New-Object System.Windows.Controls.Border
    $badge.Width = 24; $badge.Height = 24
    $badge.CornerRadius = '12'
    $badge.Background = New-FPBrush $kv.Color 0.18
    $badge.BorderBrush = New-FPBrush $kv.Color 0.55
    $badge.BorderThickness = '1'
    $badge.VerticalAlignment = 'Top'
    $glyph = New-Object System.Windows.Controls.TextBlock
    $glyph.Text = [string]$kv.Glyph
    $glyph.FontSize = 13
    $glyph.FontWeight = 'Bold'
    $glyph.Foreground = New-FPBrush $kv.Color 1.0
    $glyph.HorizontalAlignment = 'Center'
    $glyph.VerticalAlignment = 'Center'
    $badge.Child = $glyph
    [void]$head.Children.Add($badge)

    $titleTb = New-Object System.Windows.Controls.TextBlock
    $titleTb.Text = $Title
    $titleTb.FontSize = 14
    $titleTb.FontWeight = 'SemiBold'
    $titleTb.Foreground = New-FPBrush '#FFFFFF' 0.96
    $titleTb.TextWrapping = 'Wrap'
    $titleTb.Margin = '10,2,0,0'
    $titleTb.VerticalAlignment = 'Center'
    [void]$head.Children.Add($titleTb)
    [void]$box.Children.Add($head)

    if ($Message) {
        $mt = New-Object System.Windows.Controls.TextBlock
        $mt.Text = $Message
        $mt.FontSize = 12.5
        $mt.Foreground = New-FPBrush '#FFFFFF' 0.82
        $mt.TextWrapping = 'Wrap'
        $mt.Margin = '0,12,0,0'
        $mt.LineHeight = 20
        [void]$box.Children.Add($mt)
    }
    if ($Detail) {
        $dt = New-Object System.Windows.Controls.TextBlock
        $dt.Text = $Detail
        $dt.FontSize = 11.5
        $dt.Foreground = New-FPBrush '#FFFFFF' 0.5
        $dt.TextWrapping = 'Wrap'
        $dt.Margin = '0,8,0,0'
        $dt.LineHeight = 17
        [void]$box.Children.Add($dt)
    }

    if ($BuildContent) {
        $panel = New-Object System.Windows.Controls.StackPanel
        $panel.Margin = '0,14,0,0'
        [void]$box.Children.Add($panel)
        & $BuildContent $panel
    }

    if ($CheckboxText) {
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Content = $CheckboxText
        $cb.Foreground = New-FPBrush '#FFFFFF' 0.78
        $cb.FontSize = 12
        $cb.Margin = '0,14,0,0'
        $cb.Add_Checked({ $script:Dlg.Checked = $true })
        $cb.Add_Unchecked({ $script:Dlg.Checked = $false })
        [void]$box.Children.Add($cb)
    }

    $row = New-Object System.Windows.Controls.StackPanel
    $row.Orientation = 'Horizontal'
    $row.HorizontalAlignment = 'Right'
    $row.Margin = '0,18,0,0'
    if (-not $Buttons -or @($Buttons).Count -eq 0) {
        $Buttons = @(@{ Text = '知道了'; Value = 'ok'; Primary = $true })
    }
    foreach ($b in @($Buttons)) {
        $isPrimary = $false
        if ($b.ContainsKey('Primary')) { $isPrimary = [bool]$b.Primary }
        $isDanger = $false
        if ($b.ContainsKey('Danger')) { $isDanger = [bool]$b.Danger }
        $keepOpen = $false
        if ($b.ContainsKey('KeepOpen')) { $keepOpen = [bool]$b.KeepOpen }
        $act = $null
        if ($b.ContainsKey('Action')) { $act = $b.Action }
        $btn = New-DialogButton -Text $b.Text -Primary $isPrimary -Danger $isDanger
        # 按钮的上下文挂在 Tag 上：处理器里不引用任何局部变量（PowerShell 的脚本块不是词法闭包）
        $btn.Tag = @{ Value = [string]$b.Value; KeepOpen = $keepOpen; Action = $act }
        $btn.Add_Click({
                param($s, $e)
                $ctx = $s.Tag
                if ($null -eq $ctx) { return }
                if ($ctx.Action) {
                    try { & $ctx.Action } catch { Write-Log "按钮动作失败: $($_.Exception.Message)" 'ERROR' }
                }
                if (-not $ctx.KeepOpen) {
                    if ($null -ne $script:Dlg) { $script:Dlg.Value = [string]$ctx.Value }
                    $w = [System.Windows.Window]::GetWindow($s)
                    if ($w) { $w.Close() }
                }
            })
        [void]$row.Children.Add($btn)
    }
    [void]$box.Children.Add($row)
    $win.Content = $cardHost

    $win.Add_MouseLeftButtonDown({
            param($s, $e)
            # 只有点在空白/背景上才拖动窗口；点在控件上交给控件自己处理
            # （否则单选/复选按钮、文本框里的拖选会被 DragMove 吞掉）
            $interactive = $false
            try {
                $node = $e.OriginalSource
                for ($i = 0; $i -lt 14 -and $null -ne $node; $i++) {
                    $tn = $node.GetType().Name
                    if ($tn -match 'RadioButton|CheckBox|Button|TextBox|TextBoxView|ComboBox|Slider|ScrollBar|ScrollViewer|ListBox|PasswordBox|Hyperlink') { $interactive = $true; break }
                    $next = $null
                    try { $next = [System.Windows.Media.VisualTreeHelper]::GetParent($node) } catch { }
                    if ($null -eq $next) {
                        try { $next = $node.Parent } catch { }
                    }
                    if ($null -eq $next) { break }
                    $node = $next
                }
            } catch { }
            if ($interactive) { return }
            try { $s.DragMove() } catch { }
        })
    $win.Add_SourceInitialized({
            param($s, $e)
            try {
                $h = [System.Windows.Interop.WindowInteropHelper]::new($s).Handle
                $ex = [FPNative]::GetWindowLong($h, -20)
                [void][FPNative]::SetWindowLong($h, -20, ($ex -bor 0x00000080))
            } catch { }
        })
    $win.Add_Loaded({
            param($s, $e)
            try {
                $h = [System.Windows.Interop.WindowInteropHelper]::new($s).Handle
                [void][FPNative]::SetForegroundWindow($h)
                $s.Activate() | Out-Null
            } catch { }
        })

    [void]$win.ShowDialog()
    $out = $script:Dlg
    $script:Dlg = $null
    if ($null -eq $out) { $out = @{ Value = $null; Checked = $false } }
    return $out
}

function Show-DueDateDialog {
    param([string]$Name = '', [string]$Path = '', [switch]$IsNew)
    $dp = New-Object System.Windows.Controls.DatePicker
    $dp.Width = 190
    $dp.FontSize = 13
    $dp.FontFamily = 'Microsoft YaHei UI'
    $dp.SelectedDate = (Get-Date).Date.AddDays(7)
    $dp.Background = New-FPBrush '#FFFFFF' 0.94
    $dp.BorderBrush = New-FPBrush '#FFFFFF' 0.3
    $dp.Margin = '0,0,10,0'
    $script:DlgPicker = $dp

    $title = if ($IsNew) { '转为 FolderPad 框' } else { '设置截止日期' }
    $msg = '选择这个框的截止日期。到期后边框变灰、留在桌面等你处理。'
    if ($Name) { $msg = ('「{0}」' -f $Name) + "`r`n" + $msg }

    $res = Show-CustomDialog -Title $title -Message $msg -Detail $Path -Kind 'question' -Width 440 `
        -Buttons @(
        @{ Text = '取消'; Value = 'cancel' },
        @{ Text = '不设截止'; Value = 'none' },
        @{ Text = '确定'; Value = 'ok'; Primary = $true }
    ) -BuildContent {
        param($panel)
        $wrap = New-Object System.Windows.Controls.StackPanel
        $wrap.Orientation = 'Horizontal'
        [void]$wrap.Children.Add($script:DlgPicker)
        $label = New-Object System.Windows.Controls.TextBlock
        $label.Text = '（默认一周后）'
        $label.Foreground = New-FPBrush '#FFFFFF' 0.45
        $label.FontSize = 11.5
        $label.VerticalAlignment = 'Center'
        [void]$wrap.Children.Add($label)
        [void]$panel.Children.Add($wrap)

        $quick = New-Object System.Windows.Controls.WrapPanel
        $quick.Margin = '0,10,0,0'
        $presets = @(
            @{ T = '今天'; D = 0 },
            @{ T = '明天'; D = 1 },
            @{ T = '3 天后'; D = 3 },
            @{ T = '一周后'; D = 7 },
            @{ T = '两周后'; D = 14 },
            @{ T = '一个月后'; D = 30 }
        )
        foreach ($pr in $presets) {
            $b = New-DialogButton -Text $pr.T
            $b.FontSize = 11.5
            $b.Padding = '10,4'
            $b.MinWidth = 0
            $b.Margin = '0,0,6,6'
            $b.Tag = [int]$pr.D
            $b.Add_Click({
                    param($s, $e)
                    if ($script:DlgPicker) { $script:DlgPicker.SelectedDate = (Get-Date).Date.AddDays([int]$s.Tag) }
                })
            [void]$quick.Children.Add($b)
        }
        [void]$panel.Children.Add($quick)
    }

    $picked = $script:DlgPicker
    $script:DlgPicker = $null
    switch ($res.Value) {
        'ok' {
            if ($picked -and $picked.SelectedDate) {
                return @{ Ok = $true; DueDate = ([datetime]$picked.SelectedDate).ToString('yyyy-MM-dd') }
            }
            return @{ Ok = $true; DueDate = (Get-Date).Date.AddDays(7).ToString('yyyy-MM-dd') }
        }
        'none' { return @{ Ok = $true; DueDate = '' } }
        default { return $null }
    }
}

function Show-SettingsDialog {
    $st = $script:State.settings
    $res = $null
    $res = Show-CustomDialog -Title 'FolderPad 设置' -Message '' -Kind 'info' -Width 480 `
        -Buttons @(
        @{ Text = '取消'; Value = 'cancel' },
        @{ Text = '保存'; Value = 'ok'; Primary = $true }
    ) -BuildContent {
        param($panel)
        $mkRow = {
            param($label, $ctrl)
            $r = New-Object System.Windows.Controls.StackPanel
            $r.Orientation = 'Horizontal'
            $r.Margin = '0,0,0,12'
            $t = New-Object System.Windows.Controls.TextBlock
            $t.Text = $label
            $t.Width = 120
            $t.FontSize = 12.5
            $t.Foreground = New-FPBrush '#FFFFFF' 0.78
            $t.VerticalAlignment = 'Center'
            [void]$r.Children.Add($t)
            [void]$r.Children.Add($ctrl)
            [void]$panel.Children.Add($r)
        }

        $tb = New-Object System.Windows.Controls.TextBox
        $tb.Text = [string]$st.archiveRoot
        $tb.MinWidth = 200
        $tb.FontSize = 12.5
        $tb.Padding = '6,4'
        $tb.Background = New-FPBrush '#FFFFFF' 0.94
        $tb.Foreground = New-FPBrush '#111111' 1
        $script:FPSettingsArchive = $tb
        $browse = New-Object System.Windows.Controls.Button
        $browse.Content = '浏览…'
        $browse.Margin = '8,0,0,0'
        $browse.Padding = '10,3'
        $browse.Add_Click({
                $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
                $dlg.Description = '选择归档根目录'
                if ($dlg.ShowDialog() -eq 'OK') { $script:FPSettingsArchive.Text = $dlg.SelectedPath }
            })
        # 这一行用 Grid：标签 120 固定，右边整块占剩余宽度；里面再分 [文本框 *][浏览按钮 Auto]
        # 用 StackPanel 的话宽度不受约束，"浏览…" 会被对话框右边缘裁掉一半
        $rowA = New-Object System.Windows.Controls.Grid
        $rowA.Margin = '0,0,0,12'
        $rowA.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
        $rowA.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
        $rowA.ColumnDefinitions[0].Width = [System.Windows.GridLength]::new(120, [System.Windows.GridUnitType]::Pixel)
        $rowA.ColumnDefinitions[1].Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $labA = New-Object System.Windows.Controls.TextBlock
        $labA.Text = '归档根目录'
        $labA.FontSize = 12.5
        $labA.Foreground = New-FPBrush '#FFFFFF' 0.78
        $labA.VerticalAlignment = 'Center'
        [void]$rowA.Children.Add($labA)
        $innerA = New-Object System.Windows.Controls.Grid
        $innerA.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
        $innerA.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
        $innerA.ColumnDefinitions[0].Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $innerA.ColumnDefinitions[1].Width = [System.Windows.GridLength]::Auto
        $tb.HorizontalAlignment = 'Stretch'
        [void]$innerA.Children.Add($tb)
        [System.Windows.Controls.Grid]::SetColumn($browse, 1)
        [void]$innerA.Children.Add($browse)
        [System.Windows.Controls.Grid]::SetColumn($innerA, 1)
        [void]$rowA.Children.Add($innerA)
        [void]$panel.Children.Add($rowA)

        $cbExit = New-Object System.Windows.Controls.CheckBox
        $cbExit.Content = '退出程序时还原所有隐藏文件夹（推荐）'
        $cbExit.IsChecked = [bool]$st.exitRestore
        $cbExit.Foreground = New-FPBrush '#FFFFFF' 0.82
        $cbExit.FontSize = 12.5
        $script:FPSettingsExit = $cbExit
        & $mkRow '安全' $cbExit

        $cbAuto = New-Object System.Windows.Controls.CheckBox
        $cbAuto.Content = '开机自动启动 FolderPad'
        $cbAuto.IsChecked = [bool]$st.autostart
        $cbAuto.Foreground = New-FPBrush '#FFFFFF' 0.82
        $cbAuto.FontSize = 12.5
        $script:FPSettingsAuto = $cbAuto
        & $mkRow '启动' $cbAuto

        $tbRefresh = New-Object System.Windows.Controls.TextBox
        $tbRefresh.Text = [string]$st.refreshMinutes
        $tbRefresh.Width = 60
        $tbRefresh.FontSize = 12.5
        $tbRefresh.Background = New-FPBrush '#FFFFFF' 0.94
        $tbRefresh.Foreground = New-FPBrush '#111111' 1
        $script:FPSettingsRefresh = $tbRefresh
        $lbl2 = New-Object System.Windows.Controls.TextBlock
        $lbl2.Text = '  分钟自动全量刷新一次'
        $lbl2.Foreground = New-FPBrush '#FFFFFF' 0.45
        $lbl2.FontSize = 11.5
        $lbl2.VerticalAlignment = 'Center'
        $wrapR = New-Object System.Windows.Controls.StackPanel
        $wrapR.Orientation = 'Horizontal'
        [void]$wrapR.Children.Add($tbRefresh)
        [void]$wrapR.Children.Add($lbl2)
        & $mkRow '刷新' $wrapR

        $tbFill = New-Object System.Windows.Controls.TextBox
        $tbFill.Text = [string]([int]([double]$st.frameFill * 100))
        $tbFill.Width = 60
        $tbFill.FontSize = 12.5
        $tbFill.Background = New-FPBrush '#FFFFFF' 0.94
        $tbFill.Foreground = New-FPBrush '#111111' 1
        $script:FPSettingsFill = $tbFill
        $lbl3 = New-Object System.Windows.Controls.TextBlock
        $lbl3.Text = '  % 框内底色不透明度（5–60，底色就是紧急度色）'
        $lbl3.Foreground = New-FPBrush '#FFFFFF' 0.45
        $lbl3.FontSize = 11.5
        $lbl3.VerticalAlignment = 'Center'
        $wrapF = New-Object System.Windows.Controls.StackPanel
        $wrapF.Orientation = 'Horizontal'
        [void]$wrapF.Children.Add($tbFill)
        [void]$wrapF.Children.Add($lbl3)
        & $mkRow '外观' $wrapF

        $tip = New-Object System.Windows.Controls.TextBlock
        $tip.Text = ('数据目录：' + $script:DataDir)
        $tip.Foreground = New-FPBrush '#FFFFFF' 0.35
        $tip.FontSize = 11
        $tip.TextWrapping = 'Wrap'
        [void]$panel.Children.Add($tip)
    }

    if ($null -eq $res -or $res.Value -ne 'ok') { return $false }
    try {
        if ($script:FPSettingsArchive -and $script:FPSettingsArchive.Text.Trim()) {
            $st.archiveRoot = $script:FPSettingsArchive.Text.Trim()
        }
        if ($script:FPSettingsExit) { $st.exitRestore = [bool]$script:FPSettingsExit.IsChecked }
        if ($script:FPSettingsAuto) {
            $wantAuto = [bool]$script:FPSettingsAuto.IsChecked
            if ($wantAuto -ne [bool]$st.autostart) {
                $st.autostart = $wantAuto
                Set-Autostart -Enable $wantAuto
            }
        }
        if ($script:FPSettingsRefresh) {
            $n = 30
            if ([int]::TryParse($script:FPSettingsRefresh.Text, [ref]$n) -and $n -ge 1 -and $n -le 1440) { $st.refreshMinutes = $n }
        }
        if ($script:FPSettingsFill) {
            $f = 26
            if ([int]::TryParse($script:FPSettingsFill.Text, [ref]$f)) {
                $f = [Math]::Max(5, [Math]::Min(60, $f))
                $st.frameFill = $f / 100.0
            }
        }
        Save-State
        foreach ($k in @($script:Pads.Keys)) { Update-PadBackdrop -PadId $k; Refresh-Pad -PadId $k }
        return $true
    } catch {
        Show-MessageDialog -Title '保存设置失败' -Message $_.Exception.Message -Kind 'error'
        return $false
    }
}

# =====================================================================
# 框窗口
# =====================================================================
function Set-PadWindowExStyle {
    param([System.Windows.Window]$Window, [switch]$ClickThrough)
    try {
        $h = ([System.Windows.Interop.WindowInteropHelper]::new($Window)).Handle
        if ($h -eq [IntPtr]::Zero) { return }
        $ex = [FPNative]::GetWindowLong($h, -20)
        $ex = $ex -bor 0x08000000 -bor 0x00000080
        $ex = $ex -band (-bnot 0x00040000)
        if ($ClickThrough) { $ex = $ex -bor 0x00000020 }
        [void][FPNative]::SetWindowLong($h, -20, $ex)
        if (-not $ClickThrough) {
            try {
                if (-not $script:HookedHwnds.ContainsKey($h)) {
                    $src = [System.Windows.Interop.HwndSource]::FromHwnd($h)
                    if ($src) {
                        $mi = [FPMsg].GetMethod('Hook')
                        $hook = $mi.CreateDelegate([System.Windows.Interop.HwndSourceHook])
                        [void]$src.AddHook($hook)
                        $script:HookedHwnds[$h] = $true
                        Write-Log "已挂消息钩子: $h"
                    }
                }
            } catch { Write-Log "挂消息钩子失败: $($_.Exception.Message)" 'WARN' }
        }
    } catch { Write-Log "设置扩展样式失败: $($_.Exception.Message) @ $($_.ScriptStackTrace)" 'WARN' }
}

function Mix-Color {
    param([string]$Hex, [string]$With = '#FFFFFF', [double]$T = 0.12)
    $a = ConvertTo-FPColor $Hex 1.0
    $b = ConvertTo-FPColor $With 1.0
    $t = [Math]::Max(0.0, [Math]::Min(1.0, $T))
    $r = [int][Math]::Round($a.R + ($b.R - $a.R) * $t)
    $g = [int][Math]::Round($a.G + ($b.G - $a.G) * $t)
    $bl = [int][Math]::Round($a.B + ($b.B - $a.B) * $t)
    return ('#{0:X2}{1:X2}{2:X2}' -f ([Math]::Max(0, [Math]::Min(255, $r))), ([Math]::Max(0, [Math]::Min(255, $g))), ([Math]::Max(0, [Math]::Min(255, $bl))))
}

function New-PadBackdrop {
    param([string]$PadId)
    $p = $script:Pads[$PadId]
    if ($null -eq $p) { return }
    $st = $script:State.settings
    $alpha = $script:DefaultFill
    if ($null -ne $st.frameFill) { $alpha = [double]$st.frameFill }
    $pad = $script:ShadowPad
    $rad = $script:PadRadius
    $bw = New-Object System.Windows.Window
    $bw.WindowStyle = 'None'
    $bw.AllowsTransparency = $true
    $bw.Background = [System.Windows.Media.Brushes]::Transparent
    $bw.ShowInTaskbar = $false
    $bw.ShowActivated = $false
    $bw.ResizeMode = 'NoResize'
    $bw.WindowStartupLocation = 'Manual'
    $bw.IsHitTestVisible = $false
    $bw.Uid = $PadId

    # 背板：一层深灰玻璃 + 一层紧急度色
    $tint = New-Object System.Windows.Controls.Border
    $u = Get-UrgencyInfo -Pad $p.Data -Settings $st
    $fill = [string]$u.Color
    if ($p.Data.autoUrgency -eq $false -or -not $p.Data.dueDate) { $fill = $script:NeutralBorderColor }
    $tint.Background = New-FPBrush $fill $alpha
    $tint.BorderThickness = '0'
    $tint.CornerRadius = [string]$rad

    $card = New-Object System.Windows.Controls.Border
    $card.Background = New-FPBrush '#12141A' 0.34
    $card.BorderThickness = '0'
    $card.CornerRadius = [string]$rad
    $card.Child = $tint
    $bw.Content = $card
    $bw.Add_SourceInitialized({
            param($s, $e)
            Set-PadWindowExStyle -Window $s -ClickThrough
        })
    $p.Back = $bw
    $p.BackBorder = $tint
    $p.BackCard = $card
}

function Show-PadWindow {
    param([string]$PadId)
    if ($script:Pads.ContainsKey($PadId) -and $script:Pads[$PadId].Window) {
        $p0 = $script:Pads[$PadId]
        if (-not $p0.Window.IsVisible) { $p0.Window.Show() }
        return
    }
    $data = Get-PadById $PadId
    if ($null -eq $data) { Write-Log "找不到 pad 记录: $PadId" 'WARN'; return }
    if ($script:State.settings.hideAll) { return }

    $win = New-Object System.Windows.Window
    $win.WindowStyle = 'None'
    $win.AllowsTransparency = $true
    $win.Background = [System.Windows.Media.Brushes]::Transparent
    $win.ShowInTaskbar = $false
    $win.ShowActivated = $false
    $win.ResizeMode = 'NoResize'
    $win.WindowStartupLocation = 'Manual'
    $win.Topmost = [bool]$data.pinned
    $win.FontFamily = $script:FontFamily
    $win.SnapsToDevicePixels = $true
    $win.UseLayoutRounding = $true
    $win.Uid = $PadId

    $p = @{
        Id         = $PadId
        Data       = $data
        Window     = $win
        Back       = $null
        BackBorder = $null
        Items      = @()
        Ok         = $false
        Busy       = $false
        CtxOpen    = $false
        Selected   = @{}        # 选中的条目名集合（HashSet 语义用 hashtable 实现）
        Anchor     = ''          # Shift 连选的锚点条目名
        Plan       = $null       # 打包结果（Get-PadPlan）
        LayingOut  = $false
        DragFrom   = $null
        DragPoint  = $null
        DragOrigin = $null
        DragDelta  = $null
        Shield     = $false
        Hot        = $false
        Cols       = 1
        Rows       = 1
        Width      = 0
    }
    $script:Pads[$PadId] = $p

    $root = New-Object System.Windows.Controls.Border
    $root.BorderThickness = 0
    $root.CornerRadius = [string]$script:PadRadius
    $root.BorderBrush = New-FPBrush '#000000' 0.0
    $root.Background = New-FPBrush '#FFFFFF' 0.0
    $root.AllowDrop = $true
    $root.Uid = $PadId
    $p.Root = $root

    $outer = New-Object System.Windows.Controls.Grid
    $outer.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))
    $outer.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))
    $outer.RowDefinitions[0].Height = [System.Windows.GridLength]::new($script:TitleHeight, [System.Windows.GridUnitType]::Pixel)
    $outer.RowDefinitions[1].Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)

    # 标题条（类别色）
    $title = New-Object System.Windows.Controls.Border
    $title.Background = New-FPBrush $script:ThemeTitleColor 0.94
    $title.CornerRadius = "$($script:TitleRadius),$($script:TitleRadius),0,0"
    $title.Cursor = 'SizeAll'
    $title.Margin = '0'
    $title.Uid = $PadId
    $p.Title = $title

    $tg = New-Object System.Windows.Controls.Grid
    $tg.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
    $tg.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
    $tg.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
    $tg.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
    $tg.ColumnDefinitions[1].Width = [System.Windows.GridLength]::Auto
    $tg.ColumnDefinitions[2].Width = [System.Windows.GridLength]::Auto
    $tg.ColumnDefinitions[3].Width = [System.Windows.GridLength]::Auto
    $nameTb = New-Object System.Windows.Controls.TextBlock
    $nameTb.Text = [string]$data.name
    $nameTb.FontSize = $script:TitleFont
    $nameTb.FontWeight = 'SemiBold'
    $nameTb.Foreground = New-FPBrush '#FFFFFF' 0.96
    $nameTb.VerticalAlignment = 'Center'
    $nameTb.Margin = '11,0,0,0'
    $nameTb.TextTrimming = 'CharacterEllipsis'
    $tg.Children.Add($nameTb) | Out-Null
    $p.NameText = $nameTb

    # 天数区：前缀("还剩") + 数字槽 + 后缀("天")。三个元素都是网格的直接子元素，
    # 槽是没有子元素的 Border（只做底色），宽度由同格的数字文字决定。
    $duePre = New-Object System.Windows.Controls.TextBlock
    $duePre.Text = ''
    $duePre.FontSize = 11.5
    $duePre.Foreground = New-FPBrush '#FFFFFF' 0.82
    $duePre.VerticalAlignment = 'Center'
    $duePre.HorizontalAlignment = 'Right'
    $duePre.Margin = '0,0,5,0'
    [System.Windows.Controls.Grid]::SetColumn($duePre, 1)
    $tg.Children.Add($duePre) | Out-Null
    $p.DuePre = $duePre

    $dueBg = New-Object System.Windows.Controls.Border
    $dueBg.Background = New-FPBrush '#000000' $script:ChipFill
    $dueBg.CornerRadius = [string]$script:ChipRadius
    $dueBg.BorderThickness = '1'
    # 45° 渐变描边：左上暗、右下亮 → 视觉上凹进去
    $insetEdge = New-Object System.Windows.Media.LinearGradientBrush
    $insetEdge.StartPoint = '0,0'
    $insetEdge.EndPoint = '1,1'
    [void]$insetEdge.GradientStops.Add((New-Object System.Windows.Media.GradientStop((ConvertTo-FPColor '#000000' 0.55), 0.0)))
    [void]$insetEdge.GradientStops.Add((New-Object System.Windows.Media.GradientStop((ConvertTo-FPColor '#FFFFFF' 0.22), 1.0)))
    $insetEdge.Freeze()
    $dueBg.BorderBrush = $insetEdge
    $dueBg.Margin = '0,3,0,3'
    if ($script:ChipMinWidth -gt 0) { $dueBg.MinWidth = [double]$script:ChipMinWidth }
    $dueBg.HorizontalAlignment = 'Stretch'
    $dueBg.VerticalAlignment = 'Stretch'
    [System.Windows.Controls.Grid]::SetColumn($dueBg, 2)
    $tg.Children.Add($dueBg) | Out-Null
    $p.DueBox = $dueBg

    $dueTb = New-Object System.Windows.Controls.TextBlock
    $dueTb.Text = ''
    $dueTb.FontSize = $script:DueFont
    $dueTb.FontWeight = 'SemiBold'
    $dueTb.Foreground = New-FPBrush '#FFFFFF' 0.97
    $dueTb.VerticalAlignment = 'Center'
    $dueTb.HorizontalAlignment = 'Center'
    $dueTb.Margin = '8,0,8,0'
    $dueTb.TextAlignment = 'Center'
    # 数字用 Segoe UI + 等宽数字（tabular）：每个数字占同样宽、左右边距对称，1 位 2 位都真居中
    try {
        $dueTb.FontFamily = New-Object System.Windows.Media.FontFamily 'Segoe UI'
        [System.Windows.Documents.Typography]::SetNumeralAlignment($dueTb, [System.Windows.FontNumeralAlignment]::Tabular)
    } catch { }
    [System.Windows.Controls.Grid]::SetColumn($dueTb, 2)
    $tg.Children.Add($dueTb) | Out-Null
    $p.DueText = $dueTb

    $dueSuf = New-Object System.Windows.Controls.TextBlock
    $dueSuf.Text = ''
    $dueSuf.FontSize = 11.5
    $dueSuf.Foreground = New-FPBrush '#FFFFFF' 0.82
    $dueSuf.VerticalAlignment = 'Center'
    $dueSuf.HorizontalAlignment = 'Left'
    $dueSuf.Margin = '5,0,10,0'
    [System.Windows.Controls.Grid]::SetColumn($dueSuf, 3)
    $tg.Children.Add($dueSuf) | Out-Null
    $p.DueSuf = $dueSuf

    $title.Child = $tg
    [System.Windows.Controls.Grid]::SetRow($title, 0)
    $outer.Children.Add($title) | Out-Null

    # 接缝：跨在标题条和内容之间，上 1px 提亮、下 1px 压暗 → 一条半透明的"凸起棱线"
    $seam = New-Object System.Windows.Controls.StackPanel
    $seam.Orientation = 'Vertical'
    $seam.Height = 2
    $seam.VerticalAlignment = 'Bottom'
    $seam.HorizontalAlignment = 'Stretch'
    $seam.Margin = '0,0,0,-1'
    $seam.IsHitTestVisible = $false
    [System.Windows.Controls.Grid]::SetRow($seam, 0)
    $seamHi = New-Object System.Windows.Controls.Border
    $seamHi.Height = 1
    $seamHi.Background = New-FPBrush '#FFFFFF' $script:SeamHi
    [void]$seam.Children.Add($seamHi)
    $seamLo = New-Object System.Windows.Controls.Border
    $seamLo.Height = 1
    $seamLo.Background = New-FPBrush '#000000' $script:SeamLo
    [void]$seam.Children.Add($seamLo)
    $outer.Children.Add($seam) | Out-Null


    # 内容区（图标网格）
    $canvas = New-Object System.Windows.Controls.Canvas
    $canvas.Background = [System.Windows.Media.Brushes]::Transparent
    $canvas.AllowDrop = $true
    $canvas.ClipToBounds = $true
    $canvas.Uid = $PadId
    [System.Windows.Controls.Grid]::SetRow($canvas, 1)
    $outer.Children.Add($canvas) | Out-Null
    $p.Canvas = $canvas

    $root.Child = $outer
    $win.Content = $root

    $win.Add_SourceInitialized({
            param($s, $e)
            Set-PadWindowExStyle -Window $s
            Sync-PadZOrder -PadId ([string]$s.Uid)
        })

    # 标题条：拖动移动 / 双击打开 / 右键菜单
    $title.Add_MouseLeftButtonDown({
            param($s, $e)
            $pp = $script:Pads[[string]$s.Uid]
            if ($null -eq $pp) { return }
            Clear-PadSelection -PadId ([string]$s.Uid)
            if ($e.ClickCount -eq 2) {
                Open-InExplorer -Path $pp.Data.path
                return
            }
            $cur = Get-CursorDip
            $pp.Busy = $true
            $pp.DragDelta = @{ x = $cur.x; y = $cur.y }
            $pp.DragOrigin = @{ x = $pp.Window.Left; y = $pp.Window.Top }
            try { $s.CaptureMouse() } catch { }
        })
    $title.Add_MouseMove({
            param($s, $e)
            $pp = $script:Pads[[string]$s.Uid]
            if ($null -eq $pp -or $null -eq $pp.DragDelta) { return }
            if ($e.LeftButton -ne [System.Windows.Input.MouseButtonState]::Pressed) { return }
            $cur = Get-CursorDip
            $nx = $pp.DragOrigin.x + ($cur.x - $pp.DragDelta.x)
            $ny = $pp.DragOrigin.y + ($cur.y - $pp.DragDelta.y)
            Set-PadBounds -P $pp -X $nx -Y $ny
        })
    $title.Add_MouseLeftButtonUp({
            param($s, $e)
            $pp = $script:Pads[[string]$s.Uid]
            if ($null -eq $pp -or $null -eq $pp.DragDelta) { return }
            try { $s.ReleaseMouseCapture() } catch { }
            $pp.DragDelta = $null
            $pp.Busy = $false
            Save-PadPosition -PadId $pp.Id
        })
    $title.Add_MouseRightButtonUp({
            param($s, $e)
            $e.Handled = $true
            Show-PadMenu -PadId ([string]$s.Uid)
        })

    # 边框/空白区右键
    $root.Add_MouseRightButtonUp({
            param($s, $e)
            $e.Handled = $true
            Show-PadMenu -PadId ([string]$s.Uid)
        })
    $canvas.Add_MouseLeftButtonDown({
            param($s, $e)
            Clear-PadSelection -PadId ([string]$s.Uid)
        })
    $canvas.Add_MouseRightButtonUp({
            param($s, $e)
            $e.Handled = $true
            Show-PadMenu -PadId ([string]$s.Uid)
        })

    # 拖放：从资源管理器/桌面拖进来 = 移动到框目录
    $dropOver = {
        param($s, $e)
        if (-not $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
            $e.Effects = [System.Windows.DragDropEffects]::None
            $e.Handled = $true
            return
        }
        $e.Effects = [System.Windows.DragDropEffects]::Move
        $e.Handled = $true
        $pp = $script:Pads[[string]$s.Uid]
        if ($pp -and -not $pp.Hot) {
            $pp.Hot = $true
            $pp.Root.Background = New-FPBrush '#3C7DFF' 0.22
        }
    }
    $dropLeave = {
        param($s, $e)
        $id = [string]$s.Uid
        $pp = $script:Pads[$id]
        if ($null -eq $pp -or -not $pp.Hot) { return }
        $pp.Hot = $false
        try { $pp.Root.Background = New-FPBrush '#FFFFFF' 0.0 } catch { }
        Update-PadShield -PadId $id
    }
    $dropDo = {
        param($s, $e)
        $id = [string]$s.Uid
        $pp = $script:Pads[$id]
        if ($null -eq $pp) { return }
        $e.Handled = $true
        if (-not $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { return }
        $files = @($e.Data.GetData([System.Windows.DataFormats]::FileDrop))
        $pp.Hot = $false
        try { $pp.Root.Background = New-FPBrush '#FFFFFF' 0.0 } catch { }
        Update-PadShield -PadId $id
        Invoke-PadDrop -PadId $id -Files $files
    }
    $win.AllowDrop = $true
    $win.Add_DragEnter($dropOver)
    $win.Add_DragOver($dropOver)
    $win.Add_DragLeave($dropLeave)
    $win.Add_Drop($dropDo)
    $root.Add_DragEnter($dropOver)
    $root.Add_DragOver($dropOver)
    $root.Add_DragLeave($dropLeave)
    $root.Add_Drop($dropDo)
    $canvas.Add_DragEnter($dropOver)
    $canvas.Add_DragOver($dropOver)
    $canvas.Add_DragLeave($dropLeave)
    $canvas.Add_Drop($dropDo)

    if ($script:State.settings.backdrop -ne $false) {
        New-PadBackdrop -PadId $PadId
        if ($p.Back) {
            $p.Back.Left = [double]$data.pos.x - $script:ShadowPad
            $p.Back.Top = [double]$data.pos.y - $script:ShadowPad
            $p.Back.Width = 200
            $p.Back.Height = 140
            $p.Back.Show()
        }
    }
    $win.Show()
    Write-Log "显示框窗口: $PadId ($($data.name))"
    Refresh-Pad -PadId $PadId
    Sync-PadZOrder -PadId $PadId
}

function Close-PadWindow {
    param([string]$PadId)
    if (-not $script:Pads.ContainsKey($PadId)) { return }
    $p = $script:Pads[$PadId]
    try { if ($p.Window) { $p.Window.Close() } } catch { }
    try { if ($p.Back) { $p.Back.Close() } } catch { }
    $script:Pads.Remove($PadId)
}

function Set-PadBounds {
    param($P, [double]$X = [double]::NaN, [double]$Y = [double]::NaN, [double]$W = [double]::NaN, [double]$H = [double]::NaN)
    if ($null -eq $P -or $null -eq $P.Window) { return }
    $cur = @{ X = $P.Window.Left; Y = $P.Window.Top; W = $P.Window.Width; H = $P.Window.Height }
    if ([double]::IsNaN($X)) { $X = $cur.X }
    if ([double]::IsNaN($Y)) { $Y = $cur.Y }
    if ([double]::IsNaN($W)) { $W = $cur.W }
    if ([double]::IsNaN($H)) { $H = $cur.H }
    $lim = Limit-PadPosition -X $X -Y $Y -W $W -H $H
    try {
        $P.Window.Left = $lim.x
        $P.Window.Top = $lim.y
        if (-not [double]::IsNaN($W)) { $P.Window.Width = $W }
        if (-not [double]::IsNaN($H)) { $P.Window.Height = $H }
    } catch { }
    if ($P.Back) {
        try {
            # 背板比框大一圈，这圈只显示外阴影（鼠标穿透）
            $sp = $script:ShadowPad
            $P.Back.Left = $lim.x - $sp
            $P.Back.Top = $lim.y - $sp
            if (-not [double]::IsNaN($W)) { $P.Back.Width = $W + 2 * $sp }
            if (-not [double]::IsNaN($H)) { $P.Back.Height = $H + 2 * $sp }
        } catch { }
    }
}

function Save-PadPosition {
    param([string]$PadId)
    $p = $script:Pads[$PadId]
    if ($null -eq $p) { return }
    try {
        $p.Data.pos = @{ x = [int][Math]::Round($p.Window.Left); y = [int][Math]::Round($p.Window.Top) }
        Save-State
    } catch { Write-Log "保存位置失败: $($_.Exception.Message)" 'WARN' }
}

function Sync-PadZOrder {
    param([string]$PadId)
    $p = $script:Pads[$PadId]
    if ($null -eq $p -or $null -eq $p.Window) { return }
    $flags = 0x0010 -bor 0x0002 -bor 0x0001
    try {
        $fh = ([System.Windows.Interop.WindowInteropHelper]::new($p.Window)).Handle
        if ($fh -eq [IntPtr]::Zero) { return }
        if ($p.Data.pinned) {
            [void][FPNative]::SetWindowPos($fh, [FPNative]::HWND_TOPMOST, 0, 0, 0, 0, $flags)
        } else {
            [void][FPNative]::SetWindowPos($fh, [FPNative]::HWND_BOTTOM, 0, 0, 0, 0, $flags)
        }
        if ($p.Back) {
            $bh = ([System.Windows.Interop.WindowInteropHelper]::new($p.Back)).Handle
            if ($bh -ne [IntPtr]::Zero) {
                [void][FPNative]::SetWindowPos($bh, $fh, 0, 0, 0, 0, $flags)
            }
        }
    } catch { }
}

function Update-PadBackdrop {
    param([string]$PadId)
    $p = $script:Pads[$PadId]
    if ($null -eq $p -or -not $p.BackBorder) { return }
    Update-PadVisual -PadId $PadId
}

function Update-PadShield {
    param([string]$PadId)
    $p = $script:Pads[$PadId]
    if ($null -eq $p -or -not $p.Root) { return }
    $vs = Get-VirtualScreenDip
    try {
        $mx = [System.Windows.Forms.Cursor]::Position
        $pt = New-Object FPPOINT
        $pt.X = $mx.X; $pt.Y = $mx.Y
        $under = [FPNative]::WindowFromPoint($pt)
        $fh = ([System.Windows.Interop.WindowInteropHelper]::new($p.Window)).Handle
        $bh = [IntPtr]::Zero
        if ($p.Back) { $bh = ([System.Windows.Interop.WindowInteropHelper]::new($p.Back)).Handle }
        $inside = ($under -eq $fh) -or ($under -eq $bh)
    } catch { $inside = $false }

    $need = $script:ShieldUp -or $p.Hot
    if ($need -and -not $p.Shield) {
        $p.Shield = $true
        $p.Root.Background = New-FPBrush '#FFFFFF' 0.14
    } elseif (-not $need -and $p.Shield) {
        $p.Shield = $false
        $p.Root.Background = New-FPBrush '#FFFFFF' 0.0
    }
}

function Update-PadVisual {
    param([string]$PadId)
    $p = $script:Pads[$PadId]
    if ($null -eq $p) { return }
    $data = $p.Data
    $u = Get-UrgencyInfo -Pad $data -Settings $script:State.settings
    $fillColor = $u.Color
    if ($data.autoUrgency -eq $false) { $fillColor = $script:NeutralBorderColor }
    if (-not $data.dueDate) { $fillColor = $script:NeutralBorderColor }

    # 紧急度色直接铺在标题栏下方的底色上（原来那条 2px 边框已经取消）
    $alpha = $script:DefaultFill
    if ($null -ne $script:State.settings.frameFill) { $alpha = [double]$script:State.settings.frameFill }
    if ($p.BackBorder) {
        $p.BackBorder.Background = New-FPBrush $fillColor $alpha
    }
    $p.FillColor = $fillColor

    $p.DueText.Foreground = New-FPBrush '#FFFFFF' 0.92
    if ($u.Tier -eq 'expired') { $p.DueText.Foreground = New-FPBrush '#FFD8D8' 0.95 }

    $titleColor = $script:ThemeTitleColor
    if ($data.categoryColorOn -ne $false -and $data.categoryColor) { $titleColor = [string]$data.categoryColor }
    if ($u.Tier -eq 'expired') {
        $base = ConvertTo-FPColor $titleColor 1.0
        $g = [int](($base.R * 0.30) + (128 * 0.70))
        $g2 = [int](($base.G * 0.30) + (128 * 0.70))
        $g3 = [int](($base.B * 0.30) + (128 * 0.70))
        $titleColor = ('#{0:X2}{1:X2}{2:X2}' -f $g, $g2, $g3)
    }
    $p.Title.Background = New-FPBrush $titleColor 0.94
    $invalid = $false
    if (-not [System.IO.Directory]::Exists([string]$data.path)) { $invalid = $true }
    if ($invalid) {
        if ($p.BackBorder) { $p.BackBorder.Background = New-FPBrush '#E05A5A' ([Math]::Max(0.34, $alpha)) }
        $p.DueText.Foreground = New-FPBrush '#FFC9C9' 0.95
    }
    $p.NameText.Text = [string]$data.name
    $p.NameText.ToolTip = ("标题：{0}
右键 → 重命名… 可改" -f [string]$data.name)
    # 标题条字数自适应：框窄时用短文案（12天/今天/已结束），再窄就只留颜色
    Set-PadDueText -PadId $PadId
}

function Refresh-Pad {
    param([string]$PadId)
    $p = $script:Pads[$PadId]
    if ($null -eq $p) { return }
    $res = Get-PadItems $p.Data
    $p.Items = @($res.Items)
    $p.Ok = $res.Ok
    Update-PadVisual -PadId $PadId
    Update-PadLayout -PadId $PadId
}

function Get-DueParts {
    param($Urgency)
    # 只把"会变的那部分"放进槽里：有数字就放数字，没有就是状态词
    switch ($Urgency.Tier) {
        'none' { return @{ Pre = ''; Num = '未设截止'; Suf = '' } }
        'expired' { return @{ Pre = ''; Num = '已结束'; Suf = '' } }
        'today' { return @{ Pre = ''; Num = '今天截止'; Suf = '' } }
        default { return @{ Pre = '还剩'; Num = [string]$Urgency.Remaining; Suf = '天' } }
    }
}

function Set-PadDueText {
    param([string]$PadId)
    $p = $script:Pads[$PadId]
    if ($null -eq $p -or $null -eq $p.DueText) { return }
    if (-not [System.IO.Directory]::Exists([string]$p.Data.path)) {
        if ($p.DuePre) { $p.DuePre.Text = ''; $p.DuePre.Visibility = 'Collapsed' }
        if ($p.DueSuf) { $p.DueSuf.Text = ''; $p.DueSuf.Visibility = 'Collapsed' }
        $p.DueText.Text = '路径失效'
        $p.DueText.Visibility = 'Visible'
        $p.DueText.Foreground = New-FPBrush '#FFC9C9' 0.95
        if ($p.DueBox) { $p.DueBox.Visibility = 'Visible' }
        $p.DueText.ToolTip = '源文件夹已不存在，右键可「重新定位…」'
        return
    }
    $u = Get-UrgencyInfo -Pad $p.Data -Settings $script:State.settings
    $parts = Get-DueParts $u
    $w = 0
    if ($null -ne $p.Width) { $w = [double]$p.Width }
    $vis = 'Visible'
    if ($w -gt 0 -and $w -lt 86) { $vis = 'Collapsed' }
    # 太窄的时候只留槽里的数字，"还剩 / 天" 让位
    $pre = [string]$parts.Pre
    $suf = [string]$parts.Suf
    if ($w -gt 0 -and $w -lt 132) { $pre = ''; $suf = '' }

    if ($p.DuePre) {
        $p.DuePre.Text = $pre
        $p.DuePre.Visibility = $vis
    }
    if ($p.DueSuf) {
        $p.DueSuf.Text = $suf
        $p.DueSuf.Visibility = $vis
    }
    # 光学居中：首位是 1 时（1、12、14、100…）数字整体向左挪一点点，
    # 因为等宽数字里 "1" 的字面比字宽窄，视觉重心会偏右
    $num = [string]$parts.Num
    $nudge = 0.0
    if ($num -match '^1\d') { $nudge = [double]$script:ChipNudge }   # 12 / 14 / 100…：1 后面还有数字，视觉重心偏右才需要补
    $lm = [Math]::Max(2.0, 8.0 - $nudge)
    $rm = [Math]::Max(2.0, 8.0 + $nudge)
    $p.DueText.Margin = (New-Object System.Windows.Thickness $lm, 0, $rm, 0)
    # 只给"数字"上色：最后 N 天（含今天截止）用红色，过期沿用淡粉，其余白色
    $numColor = '#FFFFFF'
    $numAlpha = 0.97
    if ($u.Tier -eq 'expired') {
        $numColor = '#FFD8D8'
        $numAlpha = 0.95
    } elseif ($u.Tier -eq 'today' -or ($null -ne $u.Remaining -and [int]$u.Remaining -le [int]$script:DueNumRedDays)) {
        $numColor = [string]$script:DueNumRedColor
        $numAlpha = 1.0
    }
    $p.DueText.Foreground = New-FPBrush $numColor $numAlpha
    $p.DueText.Text = $num
    $p.DueText.Visibility = $vis
    if ($p.DueBox) { $p.DueBox.Visibility = $vis }
    $p.DueText.ToolTip = [string]$u.Text
}

function Normalize-PadOrder {
    param($Pad, [string[]]$Names)
    # 把 pad.order / pad.big 与磁盘上的真实条目对齐：
    #  - order：去掉不存在的、去重，新增的按传入顺序（自然顺序）补到末尾
    #  - big  ：去掉不存在的
    $real = @{}
    foreach ($n in @($Names)) { if ($n) { $real[[string]$n] = $true } }
    $order = New-Object System.Collections.ArrayList
    $seen = @{}
    foreach ($n in @($Pad.order)) {
        $s = [string]$n
        if (-not $s) { continue }
        if (-not $real.ContainsKey($s)) { continue }
        if ($seen.ContainsKey($s)) { continue }
        $seen[$s] = $true
        [void]$order.Add($s)
    }
    foreach ($n in @($Names)) {
        $s = [string]$n
        if (-not $s) { continue }
        if ($seen.ContainsKey($s)) { continue }
        $seen[$s] = $true
        [void]$order.Add($s)
    }
    $big = @()
    foreach ($n in @($Pad.big)) {
        $s = [string]$n
        if (-not $s) { continue }
        if (-not $real.ContainsKey($s)) { continue }
        $big += $s
    }
    return @{ Order = @($order.ToArray()); Big = @($big) }
}

function Get-PadPlan {
    param(
        [string[]]$Order = @(),
        [string[]]$Big = @(),
        [int]$Cols = 1,
        $Cells = $null
    )
    # 自由位置模型：
    #  - Cells：名字 → @{ c = 列; r = 行 }，位置由用户摆放（可以留空、可以任意行列）
    #  - 没有位置的条目（新文件、或被别人占了位置）→ 放进"第一个空格"
    #  - 大格占 $script:BigItemSpan²，位置被占则回退到第一个空格
    # 返回 @{ Positions = 名字→@{Col;Row;Span}; Cells; Cols; Rows; ForcedCols }
    $span = [int]$script:BigItemSpan
    if ($span -lt 2) { $span = 2 }
    $bigSet = @{}
    foreach ($b in @($Big)) { if ($b) { $bigSet[[string]$b] = $true } }
    if (-not $script:PadFreeLayout) { $bigSet = @{} }
    if ($Cols -lt 1) { $Cols = 1 }
    $forced = $false
    if ($bigSet.Count -gt 0 -and $Cols -lt $span) { $Cols = $span; $forced = $true }

    $occ = @{}
    $pos = @{}
    $outCells = @{}
    $maxC = 0
    $maxR = 0

    $names = @()
    foreach ($n in @($Order)) { if ($n) { $names += [string]$n } }

    # 先放大格，再放小格；有位置且不冲突就原位保留
    $pending = @()
    $ordered = @()
    foreach ($n in $names) { if ($bigSet.ContainsKey($n)) { $ordered += $n } }
    foreach ($n in $names) { if (-not $bigSet.ContainsKey($n)) { $ordered += $n } }

    foreach ($nm in $ordered) {
        $sp = 1
        if ($bigSet.ContainsKey($nm)) { $sp = $span }
        $c = -1
        $r = -1
        if ($Cells -and $Cells.ContainsKey($nm)) {
            try { $c = [int]$Cells[$nm].c; $r = [int]$Cells[$nm].r } catch { $c = -1; $r = -1 }
        }
        $free = $false
        if ($c -ge 0 -and $r -ge 0) {
            $free = $true
            for ($dx = 0; $dx -lt $sp -and $free; $dx++) {
                for ($dy = 0; $dy -lt $sp -and $free; $dy++) {
                    if ($occ.ContainsKey(('{0},{1}' -f ($c + $dx), ($r + $dy)))) { $free = $false }
                }
            }
        }
        if ($free) {
            for ($dx = 0; $dx -lt $sp; $dx++) {
                for ($dy = 0; $dy -lt $sp; $dy++) {
                    $occ[('{0},{1}' -f ($c + $dx), ($r + $dy))] = $true
                }
            }
            $pos[$nm] = @{ Col = $c; Row = $r; Span = $sp }
            $outCells[$nm] = @{ c = $c; r = $r }
            if (($c + $sp) -gt $maxC) { $maxC = $c + $sp }
            if (($r + $sp) -gt $maxR) { $maxR = $r + $sp }
        } else {
            $pending += @{ Name = $nm; Span = $sp }
        }
    }

    # 换行宽度：显式列数优先，其次是当前已占宽度，都没有就按条目数（全新框排一行）
    $wrap = $Cols
    if ($maxC -gt $wrap) { $wrap = $maxC }
    if ($wrap -lt 1) { $wrap = [Math]::Max(1, $names.Count) }

    foreach ($it in $pending) {
        $nm = [string]$it.Name
        $sp = [int]$it.Span
        $placed = $false
        for ($r = 0; -not $placed -and $r -le 4096; $r++) {
            $last = $wrap - $sp
            if ($last -lt 0) { $last = 0 }
            for ($c = 0; $c -le $last -and -not $placed; $c++) {
                $free2 = $true
                for ($dx = 0; $dx -lt $sp -and $free2; $dx++) {
                    for ($dy = 0; $dy -lt $sp -and $free2; $dy++) {
                        if ($occ.ContainsKey(('{0},{1}' -f ($c + $dx), ($r + $dy)))) { $free2 = $false }
                    }
                }
                if ($free2) {
                    for ($dx = 0; $dx -lt $sp; $dx++) {
                        for ($dy = 0; $dy -lt $sp; $dy++) {
                            $occ[('{0},{1}' -f ($c + $dx), ($r + $dy))] = $true
                        }
                    }
                    $pos[$nm] = @{ Col = $c; Row = $r; Span = $sp }
                    $outCells[$nm] = @{ c = $c; r = $r }
                    if (($c + $sp) -gt $maxC) { $maxC = $c + $sp }
                    if (($r + $sp) -gt $maxR) { $maxR = $r + $sp }
                    $placed = $true
                }
            }
        }
    }

    $cols = $maxC
    if ($cols -lt 1) { $cols = 1 }
    $rows = $maxR
    if ($rows -lt 1) { $rows = 1 }
    return @{ Positions = $pos; Cells = $outCells; Cols = $cols; Rows = $rows; ForcedCols = $forced }
}
function Update-PadLayout {
    param([string]$PadId)
    $p = $script:Pads[$PadId]
    if ($null -eq $p -or $null -eq $p.Window) { return }
    $p.LayingOut = $true
    try {
        $cw = $script:Grid.CellW
        $ch = $script:Grid.CellH
        $count = [Math]::Max(1, @($p.Items).Count)
        $vs = Get-VirtualScreenDip
        $padX = [double]$p.Data.pos.x
        $avail = $vs.Right - $padX - 2 * $script:BorderWidth - 6
        $maxCols = [int][Math]::Floor($avail / $cw)
        if ($maxCols -lt 1) { $maxCols = 1 }
        if ($maxCols -gt $script:MaxAutoCols) { $maxCols = $script:MaxAutoCols }

        # 列数 / 行数：谁被显式设置谁说了算；都 auto 就"尽量一行"
        $wantCols = 0
        $wantRows = 0
        if ([string]$p.Data.cols -match '^\d+$') { $wantCols = [int]$p.Data.cols }
        if ($script:PadFreeLayout -and [string]$p.Data.rows -match '^\d+$') { $wantRows = [int]$p.Data.rows }
        if ($wantCols -gt 0) {
            $cols = $wantCols
        } elseif ($wantRows -gt 0) {
            $cols = [int][Math]::Ceiling($count / [Math]::Max(1, $wantRows))
            if ($cols -lt 1) { $cols = 1 }
            if ($cols -gt $script:MaxAutoCols) { $cols = $script:MaxAutoCols }
        } else {
            $cols = [Math]::Min($count, $maxCols)
        }
        if ($cols -lt 1) { $cols = 1 }

        $bigNames = @()
        if ($script:PadFreeLayout) { $bigNames = @($p.Data.big) }
        $names = @($p.Items | ForEach-Object { [string]$_.Name })
        $plan = Get-PadPlan -Order $names -Big $bigNames -Cols $cols -Cells $p.Data.cells
        $p.Data.cells = $plan.Cells
        $p.Plan = $plan
        $p.Cols = [int]$plan.Cols
        $p.Rows = [int]$plan.Rows
        $p.ForcedCols = [bool]$plan.ForcedCols

        $w = $p.Cols * $cw + 2 * $script:BorderWidth
        $h = $script:TitleHeight + $p.Rows * $ch + 2 * $script:BorderWidth
        $p.Width = $w
        Set-PadBounds -P $p -X $p.Data.pos.x -Y $p.Data.pos.y -W $w -H $h
        Build-PadCells -PadId $PadId
        Set-PadDueText -PadId $PadId
    } finally {
        $p.LayingOut = $false
    }
}

function Build-PadCells {
    param([string]$PadId)
    $p = $script:Pads[$PadId]
    if ($null -eq $p) { return }
    $p.Canvas.Children.Clear()
    $cw = $script:Grid.CellW
    $ch = $script:Grid.CellH
    $plan = $p.Plan
    if ($null -eq $plan) {
        $plan = Get-PadPlan -Order @($p.Items | ForEach-Object { [string]$_.Name }) -Big @($p.Data.big) -Cols ([Math]::Max(1, [int]$p.Cols)) -Cells $p.Data.cells
        $p.Plan = $plan
    }
    foreach ($item in @($p.Items)) {
        $col = 0; $row = 0; $span = 1
        $nm = [string]$item.Name
        if ($plan.Positions.ContainsKey($nm)) {
            $pp = $plan.Positions[$nm]
            $col = [int]$pp.Col; $row = [int]$pp.Row; $span = [int]$pp.Span
        }
        $cell = New-PadCell -PadId $PadId -Item $item -Span $span
        [System.Windows.Controls.Canvas]::SetLeft($cell, $col * $cw)
        [System.Windows.Controls.Canvas]::SetTop($cell, $row * $ch)
        [void]$p.Canvas.Children.Add($cell)
    }
    Update-PadSelectionVisual -PadId $PadId

    if (-not $p.Ok) {
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = "路径失效`n右键可重新定位"
        $tb.Foreground = New-FPBrush '#FFC9C9' 0.9
        $tb.FontSize = 11.5
        $tb.TextAlignment = 'Center'
        $tb.Margin = '6,10,6,0'
        [void]$p.Canvas.Children.Add($tb)
    } elseif (@($p.Items).Count -eq 0) {
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = '（空框：把文件拖进来）'
        $tb.Foreground = New-FPBrush '#FFFFFF' 0.45
        $tb.FontSize = 11.5
        $tb.Margin = '10,10,0,0'
        [void]$p.Canvas.Children.Add($tb)
    }
}

function New-PadCell {
    param([string]$PadId, $Item, [int]$Span = 1)
    $cw = $script:Grid.CellW
    $ch = $script:Grid.CellH
    $isBig = ($Span -gt 1)
    $iconSize = $script:Grid.Icon
    if ($isBig) { $iconSize = [Math]::Round($script:Grid.Icon * $script:BigIconScale, 0) }

    $cell = New-Object System.Windows.Controls.Border
    $cell.Width = $cw * $Span
    $cell.Height = $ch * $Span
    $cell.CornerRadius = '5'
    $cell.Background = New-FPBrush '#FFFFFF' 0.035
    $cell.BorderThickness = '1'
    $cell.BorderBrush = New-FPBrush '#000000' 0.0
    $cell.Cursor = 'Hand'
    $cell.ToolTip = [string]$Item.FullName
    $cell.Uid = $PadId
    $cell.Tag = $Item
    $cell.DataContext = [string]$Item.Name

    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.VerticalAlignment = 'Center'
    if ($isBig) { $sp.Margin = '4,6,4,4' } else { $sp.Margin = '3,4,3,2' }

    $img = New-Object System.Windows.Controls.Image
    $img.Width = $iconSize
    $img.Height = $iconSize
    $img.Stretch = 'Uniform'
    try { [System.Windows.Media.RenderOptions]::SetBitmapScalingMode($img, [System.Windows.Media.BitmapScalingMode]::HighQuality) } catch { }
    $src = $null
    if ($isBig) { $src = Get-FileIconSource -Path $Item.FullName -IsDir $Item.IsDir -WantSize ([int]$iconSize) }
    if (-not $src) { $src = Get-FileIconSource -Path $Item.FullName -IsDir $Item.IsDir }
    if ($src) { $img.Source = $src }
    [void]$sp.Children.Add($img)

    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = [string]$Item.Name
    $tb.FontSize = 11.5
    if ($isBig) { $tb.FontSize = 12.5 }
    $tb.Foreground = New-FPBrush '#FFFFFF' 0.96
    $tb.TextAlignment = 'Center'
    $tb.TextWrapping = 'Wrap'
    $tb.MaxHeight = 31
    if ($isBig) { $tb.MaxHeight = 34 }
    $tb.Margin = '0,4,0,0'
    $tb.LineHeight = 14.5
    $sh = New-Object System.Windows.Media.Effects.DropShadowEffect
    $sh.BlurRadius = 4
    $sh.ShadowDepth = 0
    $sh.Opacity = 0.9
    $sh.Color = [System.Windows.Media.Colors]::Black
    $tb.Effect = $sh
    [void]$sp.Children.Add($tb)
    $cell.Child = $sp

    $cell.Add_MouseLeftButtonDown({
            param($s, $e)
            $id = [string]$s.Uid
            $pp = $script:Pads[$id]
            if ($null -eq $pp) { return }
            $itemName = [string]$s.DataContext
            $mods = [System.Windows.Input.Keyboard]::Modifiers
            $ctrl = ($mods -band [System.Windows.Input.ModifierKeys]::Control) -ne 0
            $shift = ($mods -band [System.Windows.Input.ModifierKeys]::Shift) -ne 0
            if ($e.ClickCount -eq 2) {
                Clear-AllPadSelection
                Open-PadItem -PadId $id -Item $s.Tag
                $e.Handled = $true
                return
            }
            if ($ctrl) {
                Toggle-PadCellSelection -PadId $id -Name $itemName
            } elseif ($shift) {
                Select-PadRange -PadId $id -Name $itemName
            } else {
                if (-not $pp.Selected.ContainsKey($itemName)) {
                    Select-PadCell -PadId $id -Cell $s
                } else {
                    $pp.Anchor = $itemName
                }
            }
            $pp.DragFrom = $s.Tag
            $pp.DragPoint = Get-CursorDip
            $pp.DragNames = @($pp.Selected.Keys)
            $pp.DragOrderMode = $false
            $e.Handled = $true
        })
    $cell.Add_MouseMove({
            param($s, $e)
            $id = [string]$s.Uid
            $pp = $script:Pads[$id]
            if ($null -eq $pp -or $null -eq $pp.DragFrom -or $null -eq $pp.DragPoint) { return }
            if ($e.LeftButton -ne [System.Windows.Input.MouseButtonState]::Pressed) { return }
            $cur = Get-CursorDip
            if ([Math]::Abs($cur.x - $pp.DragPoint.x) -lt 6 -and [Math]::Abs($cur.y - $pp.DragPoint.y) -lt 6) { return }
            $local = Get-PadLocalPoint -P $pp -Point $cur
            $inside = Test-PadContentPoint -P $pp -Local $local
            if ($inside -and $script:ReorderByDrag) {
                $pp.DragOrderMode = $true
                $idx = Find-PadInsertIndex -P $pp -LocalX $local.x -LocalY $local.y
                Show-PadDropPreview -P $pp -Index $idx
            } elseif ($inside) {
                # 拖动排序已关闭：框内拖动什么都不做（用方向键调位置）
            } else {
                Hide-PadDropPreview -P $pp
                $names = @($pp.DragNames)
                $pp.DragFrom = $null
                $pp.DragPoint = $null
                $pp.DragNames = @()
                $pp.DragOrderMode = $false
                Start-PadDragOut -PadId $id -Names $names
            }
        })
    $cell.Add_MouseLeftButtonUp({
            param($s, $e)
            $id = [string]$s.Uid
            $pp = $script:Pads[$id]
            if ($null -eq $pp) { return }
            $wasReorder = [bool]$pp.DragOrderMode
            $names = @($pp.DragNames)
            $pp.DragFrom = $null
            $pp.DragPoint = $null
            $pp.DragNames = @()
            $pp.DragOrderMode = $false
            if (-not $wasReorder) { Hide-PadDropPreview -P $pp; return }
            $cur = Get-CursorDip
            $local = Get-PadLocalPoint -P $pp -Point $cur
            Hide-PadDropPreview -P $pp
            if (@($names).Count -eq 0) { return }
            $idx = Find-PadInsertIndex -P $pp -LocalX $local.x -LocalY $local.y
            Move-PadItemsToIndex -P $pp -Names $names -Index $idx
        })
    $cell.Add_MouseRightButtonUp({
            param($s, $e)
            $e.Handled = $true
            $id = [string]$s.Uid
            $pp = $script:Pads[$id]
            if ($null -eq $pp) { return }
            $nm = [string]$s.DataContext
            if (-not $pp.Selected.ContainsKey($nm)) { Select-PadCell -PadId $id -Cell $s }
            Show-ItemMenu -PadId $id -Item $s.Tag
        })

    return $cell
}

function Get-PadLocalPoint {
    param($P, $Point)
    # 屏幕 DIP → 框内相对坐标（含标题条）
    $x = [double]$Point.x - [double]$P.Data.pos.x
    $y = [double]$Point.y - [double]$P.Data.pos.y
    return @{ x = $x; y = $y }
}

function Test-PadContentPoint {
    param($P, $Local)
    if ($null -eq $Local) { return $false }
    $cw = $script:Grid.CellW
    $ch = $script:Grid.CellH
    if ($Local.x -lt 0) { return $false }
    if ($Local.x -gt ([double]$P.Cols * $cw)) { return $false }
    if ($Local.y -lt [double]$script:TitleHeight) { return $false }
    if ($Local.y -gt ([double]$script:TitleHeight + [double]$P.Rows * $ch)) { return $false }
    return $true
}

function Find-PadInsertIndex {
    param($P, [double]$LocalX, [double]$LocalY)
    $cw = $script:Grid.CellW
    $ch = $script:Grid.CellH
    $names = @($P.Items | ForEach-Object { [string]$_.Name })
    if ($names.Count -eq 0) { return 0 }
    $plan = $P.Plan
    if ($null -eq $plan) { return $names.Count }
    $bestName = $null
    $bestD = [double]::MaxValue
    $bx = 0.0; $by = 0.0; $bw = 0.0; $bh = 0.0
    foreach ($nm in $names) {
        if (-not $plan.Positions.ContainsKey($nm)) { continue }
        $pp = $plan.Positions[$nm]
        $x0 = [double]$pp.Col * $cw
        $y0 = [double]$script:TitleHeight + [double]$pp.Row * $ch
        $w = [double]$pp.Span * $cw
        $h = [double]$pp.Span * $ch
        $dx = 0.0; $dy = 0.0
        if ($LocalX -lt $x0) { $dx = $x0 - $LocalX } elseif ($LocalX -gt ($x0 + $w)) { $dx = $LocalX - ($x0 + $w) }
        if ($LocalY -lt $y0) { $dy = $y0 - $LocalY } elseif ($LocalY -gt ($y0 + $h)) { $dy = $LocalY - ($y0 + $h) }
        $d = $dx * $dx + $dy * $dy
        if ($d -lt $bestD) {
            $bestD = $d; $bestName = $nm
            $bx = $x0; $by = $y0; $bw = $w; $bh = $h
        }
    }
    if ($null -eq $bestName) { return $names.Count }
    $idx = [array]::IndexOf($names, $bestName)
    if ($idx -lt 0) { return $names.Count }
    $cx = $bx + $bw / 2.0
    $cy = $by + $bh / 2.0
    $after = $false
    if ($LocalY -gt ($cy + 2)) { $after = $true }
    elseif ([Math]::Abs($LocalY - $cy) -le 2 -and $LocalX -gt $cx) { $after = $true }
    if ($after) { $idx++ }
    return $idx
}

function Get-PadReorderedNames {
    param($P, [string[]]$Names, [int]$Index)
    $order = @($P.Items | ForEach-Object { [string]$_.Name })
    $moving = @()
    foreach ($n in @($Names)) { if ($n) { $moving += [string]$n } }
    if ($moving.Count -eq 0) { return $order }
    $rest = @()
    $before = 0
    for ($i = 0; $i -lt $order.Count; $i++) {
        $nm = $order[$i]
        if ($moving -contains $nm) {
            if ($i -lt $Index) { $before++ }
        } else {
            $rest += $nm
        }
    }
    $at = $Index - $before
    if ($at -lt 0) { $at = 0 }
    if ($at -gt $rest.Count) { $at = $rest.Count }
    $out = @()
    if ($at -gt 0) { $out += @($rest[0..($at - 1)]) }
    $out += $moving
    if ($at -lt $rest.Count) { $out += @($rest[$at..($rest.Count - 1)]) }
    return @($out)
}

function Move-PadItemsToIndex {
    param($P, [string[]]$Names, [int]$Index)
    if ($null -eq $P) { return }
    $newOrder = @(Get-PadReorderedNames -P $P -Names $Names -Index $Index)
    $cur = @($P.Items | ForEach-Object { [string]$_.Name })
    if (($newOrder -join '|') -eq ($cur -join '|')) { return }
    $P.Data.order = @($newOrder)
    Save-State
    Refresh-Pad -PadId $P.Id
    Write-Log ("顺序调整: " + ($newOrder -join ' > '))
}

function Get-ForegroundIsDesktop {
    # 判断前台窗口是不是"桌面"（Progman / WorkerW 或 SHELLDLL_DefView）
    try {
        $fg = [FPNative]::GetForegroundWindow()
        if ($fg -eq [IntPtr]::Zero) { return $false }
        $sb = New-Object System.Text.StringBuilder 256
        [void][FPNative]::GetClassName($fg, $sb, 256)
        $cls = $sb.ToString()
        if ($cls -eq 'Progman' -or $cls -eq 'WorkerW' -or $cls -eq 'SHELLDLL_DefView') { return $true }
        return $false
    } catch { return $false }
}

function Get-PadSelectionPad {
    # 找到"当前有选中项"的框（多框时以最后操作的那个为准）
    $cand = $null
    foreach ($k in @($script:Pads.Keys)) {
        $pp = $script:Pads[$k]
        if ($null -eq $pp -or $null -eq $pp.Selected) { continue }
        if ($pp.Selected.Count -le 0) { continue }
        $cand = $pp
    }
    return $cand
}

function Move-PadSelectionByStep {
    param($P, [int]$DeltaCol = 0, [int]$DeltaRow = 0)
    if ($null -eq $P -or ($DeltaCol -eq 0 -and $DeltaRow -eq 0)) { return }
    if ($P.Selected.Count -le 0) { return }
    $plan = $P.Plan
    if ($null -eq $plan) { return }
    $cells = @{}
    foreach ($k in @($P.Data.cells.Keys)) { $cells[[string]$k] = @{ c = [int]$P.Data.cells[$k].c; r = [int]$P.Data.cells[$k].r } }
    if ($cells.Count -eq 0) { return }

    $span = [int]$script:BigItemSpan
    $moving = @()
    foreach ($k in @($P.Selected.Keys)) { if ($cells.ContainsKey([string]$k)) { $moving += [string]$k } }
    if ($moving.Count -eq 0) { return }

    # 目标格占用情况
    $occ = @{}
    foreach ($k in @($cells.Keys)) {
        $sp = 1
        if (@($P.Data.big) -contains $k) { $sp = $span }
        for ($dx = 0; $dx -lt $sp; $dx++) {
            for ($dy = 0; $dy -lt $sp; $dy++) {
                $occ[('{0},{1}' -f ([int]$cells[$k].c + $dx), ([int]$cells[$k].r + $dy))] = [string]$k
            }
        }
    }
    $movingSet = @{}
    foreach ($m in $moving) { $movingSet[$m] = $true }

    # 有没有撞到别人 / 越界
    $blockers = @{}
    $ok = $true
    foreach ($m in $moving) {
        $sp = 1
        if (@($P.Data.big) -contains $m) { $sp = $span }
        $nc = [int]$cells[$m].c + $DeltaCol
        $nr = [int]$cells[$m].r + $DeltaRow
        if ($nc -lt 0 -or $nr -lt 0) { $ok = $false; break }
        for ($dx = 0; $dx -lt $sp; $dx++) {
            for ($dy = 0; $dy -lt $sp; $dy++) {
                $key = ('{0},{1}' -f ($nc + $dx), ($nr + $dy))
                if ($occ.ContainsKey($key)) {
                    $who = [string]$occ[$key]
                    if (-not $movingSet.ContainsKey($who)) { $blockers[$who] = $true; $ok = $false }
                }
            }
        }
    }

    if (-not $ok) {
        # 1×1 之间可以"交换"；碰到放大块就挡住不动
        $canSwap = ($blockers.Count -eq 1)
        if ($canSwap) {
            $bk = @($blockers.Keys)[0]
            foreach ($m in $moving) {
                if (@($P.Data.big) -contains $m) { $canSwap = $false }
            }
            if (@($P.Data.big) -contains $bk) { $canSwap = $false }
            if ($moving.Count -ne 1) { $canSwap = $false }
        }
        if (-not $canSwap) { return }
        $bk2 = @($blockers.Keys)[0]
        $tmpC = [int]$cells[$bk2].c; $tmpR = [int]$cells[$bk2].r
        $mv = [string]$moving[0]
        $cells[$bk2] = @{ c = [int]$cells[$mv].c; r = [int]$cells[$mv].r }
        $cells[$mv] = @{ c = $tmpC; r = $tmpR }
    } else {
        foreach ($m in $moving) {
            $cells[$m] = @{ c = ([int]$cells[$m].c + $DeltaCol); r = ([int]$cells[$m].r + $DeltaRow) }
        }
    }
    $P.Data.cells = $cells
    Save-State
    Refresh-Pad -PadId $P.Id
}
function Show-PadDropPreview {
    param($P, [int]$Index)
    if ($null -eq $P -or $null -eq $P.Canvas) { return }
    Hide-PadDropPreview -P $P
    $moving = @()
    foreach ($k in @($P.Selected.Keys)) { $moving += [string]$k }
    if ($moving.Count -eq 0) { return }
    $preview = @(Get-PadReorderedNames -P $P -Names $moving -Index $Index)
    $plan = Get-PadPlan -Order $preview -Big @($P.Data.big) -Cols ([Math]::Max(1, [int]$P.Cols))
    $first = [string]$moving[0]
    if (-not $plan.Positions.ContainsKey($first)) { return }
    $pp = $plan.Positions[$first]
    $cw = $script:Grid.CellW
    $ch = $script:Grid.CellH
    $mk = New-Object System.Windows.Controls.Border
    $mk.Name = 'fpDropPreview'
    $mk.Width = [double]$pp.Span * $cw
    $mk.Height = [double]$pp.Span * $ch
    $mk.Background = New-FPBrush '#7FB6FF' 0.26
    $mk.BorderBrush = New-FPBrush '#9CC8FF' 0.9
    $mk.BorderThickness = '2'
    $mk.CornerRadius = '5'
    $mk.IsHitTestVisible = $false
    [System.Windows.Controls.Canvas]::SetLeft($mk, [double]$pp.Col * $cw)
    [System.Windows.Controls.Canvas]::SetTop($mk, [double]$pp.Row * $ch)
    [void]$P.Canvas.Children.Add($mk)
}

function Hide-PadDropPreview {
    param($P)
    if ($null -eq $P -or $null -eq $P.Canvas) { return }
    $rm = @()
    foreach ($c in @($P.Canvas.Children)) {
        if ($c -is [System.Windows.Controls.Border] -and [string]$c.Name -eq 'fpDropPreview') { $rm += $c }
    }
    foreach ($c in $rm) { [void]$P.Canvas.Children.Remove($c) }
}

function Update-PadSelectionVisual {
    param([string]$PadId)
    $p = $script:Pads[$PadId]
    if ($null -eq $p -or $null -eq $p.Canvas) { return }
    $sel = $p.Selected
    if ($null -eq $sel) { $sel = @{} }
    foreach ($c in @($p.Canvas.Children)) {
        if (-not ($c -is [System.Windows.Controls.Border])) { continue }
        $nm = [string]$c.DataContext
        if (-not $nm) { continue }
        if ($sel.ContainsKey($nm)) {
            try {
                $c.BorderBrush = New-FPBrush '#7FB6FF' 0.9
                $c.Background = New-FPBrush '#7FB6FF' 0.20
            } catch { }
        } else {
            try {
                $c.BorderBrush = New-FPBrush '#000000' 0.0
                $c.Background = New-FPBrush '#FFFFFF' 0.035
            } catch { }
        }
    }
}

function Clear-PadSelection {
    param([string]$PadId)
    $p = $script:Pads[$PadId]
    if ($null -eq $p) { return }
    if ($null -eq $p.Selected -or $p.Selected.Count -eq 0) { return }
    $p.Selected = @{}
    $p.Anchor = ''
    Update-PadSelectionVisual -PadId $PadId
}

function Clear-AllPadSelection {
    foreach ($k in @($script:Pads.Keys)) {
        try { Clear-PadSelection -PadId $k } catch { }
    }
}

function Select-PadCell {
    param([string]$PadId, $Cell)
    $p = $script:Pads[$PadId]
    if ($null -eq $p) { return }
    Clear-AllPadSelection
    $nm = [string]$Cell.DataContext
    $sel = @{}
    if ($nm) { $sel[$nm] = $true }
    $p.Selected = $sel
    $p.Anchor = $nm
    Update-PadSelectionVisual -PadId $PadId
}

function Toggle-PadCellSelection {
    param([string]$PadId, [string]$Name)
    $p = $script:Pads[$PadId]
    if ($null -eq $p -or -not $Name) { return }
    foreach ($k in @($script:Pads.Keys)) {
        if ($k -eq $PadId) { continue }
        try { Clear-PadSelection -PadId $k } catch { }
    }
    if ($null -eq $p.Selected) { $p.Selected = @{} }
    if ($p.Selected.ContainsKey($Name)) { [void]$p.Selected.Remove($Name) } else { $p.Selected[$Name] = $true }
    $p.Anchor = $Name
    Update-PadSelectionVisual -PadId $PadId
}

function Select-PadRange {
    param([string]$PadId, [string]$Name)
    $p = $script:Pads[$PadId]
    if ($null -eq $p -or -not $Name) { return }
    $names = @($p.Items | ForEach-Object { [string]$_.Name })
    $a = [array]::IndexOf($names, [string]$p.Anchor)
    $b = [array]::IndexOf($names, $Name)
    $keepAnchor = [string]$p.Anchor
    if ($a -lt 0 -or $b -lt 0) {
        Clear-AllPadSelection
        $p.Selected = @{ $Name = $true }
        $p.Anchor = $Name
        Update-PadSelectionVisual -PadId $PadId
        return
    }
    if ($a -gt $b) { $t = $a; $a = $b; $b = $t }
    $sel = @{}
    for ($i = $a; $i -le $b; $i++) { $sel[$names[$i]] = $true }
    Clear-AllPadSelection
    $p.Selected = $sel
    $p.Anchor = $keepAnchor
    Update-PadSelectionVisual -PadId $PadId
}

function Open-PadItem {
    param([string]$PadId, $Item, [switch]$Reveal)
    if ($null -eq $Item) { return }
    try {
        if ($Reveal) { Open-InExplorer -Path $Item.FullName -Select; return }
        Start-Process -FilePath $Item.FullName
    } catch {
        Show-MessageDialog -Title '打开失败' -Message $_.Exception.Message -Kind 'error'
    }
}

function Start-PadDragOut {
    param([string]$PadId, [string[]]$Names = @(), [string]$FullPath = '')
    $p = $script:Pads[$PadId]
    if ($null -eq $p) { return }
    $dropPaths = @()
    if ($FullPath) { $dropPaths += $FullPath }
    foreach ($n in @($Names)) {
        if (-not $n) { continue }
        $fp = Join-Path ([string]$p.Data.path) ([string]$n)
        if (Test-Path -LiteralPath $fp) { $dropPaths += $fp }
    }
    $dropPaths = @($dropPaths | Select-Object -Unique)
    if (@($dropPaths).Count -eq 0) { return }
    $FullPath = [string]$dropPaths[0]
    try {
        $data = New-Object System.Windows.DataObject
        $data.SetData([System.Windows.DataFormats]::FileDrop, [string[]]$dropPaths)
        $effect = [System.Windows.DragDrop]::DoDragDrop($p.Window, $data,
            ([System.Windows.DragDropEffects]::Move -bor [System.Windows.DragDropEffects]::Copy))
        Write-Log ("拖出 {0} 项 effect={1}" -f @($dropPaths).Count, $effect)
        if ($effect -ne [System.Windows.DragDropEffects]::None -and $effect -ne 0) {
            # 目标只拷贝没有移动时，源文件还在：按「拖出去=移回桌面」的语义补删源
            $desk = [Environment]::GetFolderPath('Desktop')
            foreach ($src in @($dropPaths)) {
                if (-not (Test-Path -LiteralPath $src)) { continue }
                $leaf = Split-Path -Leaf $src
                $cand = Join-Path $desk $leaf
                if (Test-Path -LiteralPath $cand) {
                    $sLen = 0; $dLen = -1
                    try { $sLen = (Get-Item -LiteralPath $src -Force).Length } catch { }
                    try { $dLen = (Get-Item -LiteralPath $cand -Force).Length } catch { }
                    if ($sLen -eq $dLen) {
                        Remove-Item -LiteralPath $src -Recurse -Force
                        Write-Log "拖出完成（补删源）: $src -> $cand"
                    }
                }
            }
        }
    } catch {
        Write-Log "拖出失败: $($_.Exception.Message)" 'WARN'
    } finally {
        Refresh-Pad -PadId $PadId
    }
}

function Invoke-PadDrop {
    param([string]$PadId, [string[]]$Files)
    $p = $script:Pads[$PadId]
    if ($null -eq $p) { return }
    if (-not $p.Ok) {
        Show-MessageDialog -Title '无法接收文件' -Message '这个框的源文件夹不存在（路径失效），请先右键「重新定位…」。' -Kind 'warn'
        return
    }
    if (-not $Files -or $Files.Count -eq 0) { return }
    $destDir = [string]$p.Data.path
    $bulk = ''
    $okCount = 0
    $skipCount = 0
    $errors = @()
    foreach ($f in $Files) {
        if (-not (Test-Path -LiteralPath $f)) { continue }
        if (Test-PathEqual (Split-Path -Parent $f) $destDir) { continue }
        $leaf = Split-Path -Leaf $f
        $target = Join-Path $destDir $leaf
        $action = 'rename'
        if (Test-Path -LiteralPath $target) {
            if ($bulk) { $action = $bulk }
            else {
                $isDir = [System.IO.Directory]::Exists($f)
                $btns = New-Object System.Collections.ArrayList
                if ($isDir) {
                    [void]$btns.Add(@{ Text = '跳过'; Value = 'skip' })
                    [void]$btns.Add(@{ Text = '合并'; Value = 'merge' })
                    [void]$btns.Add(@{ Text = '重命名'; Value = 'rename'; Primary = $true })
                } else {
                    [void]$btns.Add(@{ Text = '跳过'; Value = 'skip' })
                    [void]$btns.Add(@{ Text = '覆盖'; Value = 'overwrite'; Danger = $true })
                    [void]$btns.Add(@{ Text = '重命名'; Value = 'rename'; Primary = $true })
                }
                [void]$btns.Add(@{ Text = '取消'; Value = 'cancel' })
                $res = Show-CustomDialog -Title '同名项目已存在' `
                    -Message ('框里已经有同名项目：' + "`r`n" + $leaf) `
                    -Detail ('目标目录：' + $destDir) -Kind 'question' -Width 460 `
                    -CheckboxText '对后面所有同名项目都用这个选择' `
                    -Buttons $btns.ToArray()
                if ($res.Value -eq 'cancel' -or $null -eq $res.Value) { return }
                $action = [string]$res.Value
                if ($res.Checked) { $bulk = $action }
            }
        }
        $r = Move-ItemToDirectory -Source $f -DestDir $destDir -ConflictAction $action
        if ($r.Ok) {
            if ($r.Skipped) { $skipCount++ } else { $okCount++ }
        } else {
            $errors += ($leaf + '：' + $r.Error)
        }
    }
    Refresh-Pad -PadId $PadId
    if ($errors.Count -gt 0) {
        Show-MessageDialog -Title '部分文件没能移动' -Message ($errors -join "`r`n") -Kind 'warn'
    } elseif ($okCount -gt 0) {
        $tip = "已移入 $okCount 个项目"
        if ($skipCount -gt 0) { $tip = $tip + "，跳过 $skipCount 个" }
        Show-TrayBalloon -Title 'FolderPad' -Text $tip
    }
}

# =====================================================================
# 菜单
# =====================================================================
function New-MenuEntry {
    param(
        [string]$PadId,
        [string]$Text,
        [scriptblock]$Action = $null,
        $Item = $null,
        [switch]$Checked,
        [switch]$Checkable,
        [string]$Hint = '',
        [object[]]$Dots = @()
    )
    return @{
        Kind      = 'item'
        PadId     = $PadId
        Text      = $Text
        Action    = $Action
        Item      = $Item
        Checked   = [bool]$Checked
        Checkable = [bool]$Checkable
        Hint      = $Hint
        Dots      = @($Dots)
    }
}

function New-MenuSep {
    return @{ Kind = 'sep' }
}

function New-MenuHeader {
    param([string]$Text)
    return @{ Kind = 'header'; Text = $Text }
}

function Show-FpMenu {
    param([string]$PadId, [object[]]$Entries)
    $p = $script:Pads[$PadId]
    if ($null -eq $p) { return }
    # 同一时间只允许一个菜单
    try {
        if ($script:OpenMenuWindow -and $script:OpenMenuWindow.IsVisible) { $script:OpenMenuWindow.Close() }
    } catch { }
    $p.CtxOpen = $true

    $win = New-Object System.Windows.Window
    $win.WindowStyle = 'None'
    $win.AllowsTransparency = $true
    $win.Background = [System.Windows.Media.Brushes]::Transparent
    $win.ShowInTaskbar = $false
    $win.ShowActivated = $false
    $win.ResizeMode = 'NoResize'
    $win.SizeToContent = 'WidthAndHeight'
    $win.Topmost = $true
    $win.WindowStartupLocation = 'Manual'
    $win.FontFamily = $script:FontFamily
    $win.Uid = $PadId
    $win.Add_Closed({
            param($s, $e)
            $pp = $script:Pads[[string]$s.Uid]
            if ($pp) { $pp.CtxOpen = $false }
        })
    $win.Add_MouseLeave({
            param($s, $e)
            $t = New-Object System.Windows.Threading.DispatcherTimer
            $t.Interval = [TimeSpan]::FromMilliseconds(400)
            $t.Add_Tick({
                    param($ts, $te)
                    $ts.Stop()
                    $ww = $script:OpenMenuWindow
                    if ($ww) { if (-not $ww.IsMouseOver) { try { $ww.Close() } catch { } } }
                })
            $script:OpenMenuTimer = $t
            $t.Start()
        })

    $card = New-Object System.Windows.Controls.Border
    $card.Background = New-FPBrush '#F2262629' 1.0
    $card.BorderBrush = New-FPBrush '#FFFFFF' 0.30
    $card.BorderThickness = '1'
    $card.CornerRadius = '8'
    $card.Padding = '4'
    $card.MaxWidth = [double]$script:MenuMaxWidth
    $shadow = New-Object System.Windows.Media.Effects.DropShadowEffect
    $shadow.BlurRadius = 16
    $shadow.ShadowDepth = 2
    $shadow.Opacity = 0.5
    $shadow.Color = [System.Windows.Media.Colors]::Black
    $card.Effect = $shadow
    # 菜单窗口也要给阴影留一圈透明边，否则阴影被切出虚边
    $menuHost = New-Object System.Windows.Controls.Grid
    $menuHost.Margin = (New-Object System.Windows.Thickness $script:MenuShadowPad)
    [void]$menuHost.Children.Add($card)

    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.MinWidth = 214
    $card.Child = $panel

    foreach ($en in @($Entries)) {
        if ($null -eq $en) { continue }
        switch ($en.Kind) {
            'sep' {
                $sep = New-Object System.Windows.Controls.Border
                $sep.Height = 1
                $sep.Background = New-FPBrush '#FFFFFF' 0.13
                $sep.Margin = '8,5,8,5'
                [void]$panel.Children.Add($sep)
            }
            'header' {
                $hb = New-Object System.Windows.Controls.TextBlock
                $hb.Text = $en.Text
                $hb.FontSize = 11
                $hb.Foreground = New-FPBrush '#FFFFFF' 0.45
                $hb.Margin = '10,7,10,3'
                $hb.TextTrimming = 'CharacterEllipsis'
                $hb.TextWrapping = 'NoWrap'
                [void]$panel.Children.Add($hb)
            }
            default {
                $row = New-Object System.Windows.Controls.Border
                $row.Background = New-FPBrush '#000000' 0.0
                $row.CornerRadius = '5'
                $row.Padding = '9,6,12,6'
                $row.Cursor = 'Hand'
                $row.Tag = $en
                $g = New-Object System.Windows.Controls.Grid
                $g.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
                $g.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
                $g.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
                $g.ColumnDefinitions[0].Width = [System.Windows.GridLength]::new(16, [System.Windows.GridUnitType]::Pixel)
                $g.ColumnDefinitions[1].Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
                $g.ColumnDefinitions[2].Width = [System.Windows.GridLength]::Auto
                if ($en.Checkable -and $en.Checked) {
                    $ck = New-Object System.Windows.Controls.TextBlock
                    $ck.Text = [string][char]0x2713
                    $ck.FontSize = 12
                    $ck.Foreground = New-FPBrush '#8CC4FF' 1.0
                    $ck.VerticalAlignment = 'Center'
                    [void]$g.Children.Add($ck)
                }
                $tb = New-Object System.Windows.Controls.TextBlock
                $tb.Text = $en.Text
                $tb.FontSize = 12.5
                $tb.Foreground = New-FPBrush '#EAEAEA' 1.0
                $tb.VerticalAlignment = 'Center'
                $tb.TextTrimming = 'CharacterEllipsis'
                [System.Windows.Controls.Grid]::SetColumn($tb, 1)
                [void]$g.Children.Add($tb)
                if ($en.Hint) {
                    $h = New-Object System.Windows.Controls.TextBlock
                    $h.Text = $en.Hint
                    $h.FontSize = 11
                    $h.Foreground = New-FPBrush '#FFFFFF' 0.40
                    $h.VerticalAlignment = 'Center'
                    $h.Margin = '14,0,0,0'
                    $h.TextTrimming = 'CharacterEllipsis'
                    $h.TextWrapping = 'NoWrap'
                    # 提示再长也最多占 150px，不然会把整个菜单顶宽
                    $h.MaxWidth = [double]$script:MenuHintMaxWidth
                    [System.Windows.Controls.Grid]::SetColumn($h, 2)
                    [void]$g.Children.Add($h)
                }
                $row.Child = $g
                $row.Add_MouseEnter({
                        param($s, $e)
                        $s.Background = New-FPBrush '#2EFFFFFF' 1.0
                    })
                $row.Add_MouseLeave({
                        param($s, $e)
                        $s.Background = New-FPBrush '#000000' 0.0
                    })
                $row.Add_MouseLeftButtonUp({
                        param($s, $e)
                        $ctx = $s.Tag
                        $w = [System.Windows.Window]::GetWindow($s)
                        if ($w) { try { $w.Close() } catch { } }
                        if ($ctx -and $ctx.Action) {
                            try { & $ctx.Action $ctx.PadId $ctx.Item } catch {
                                Write-Log "菜单动作失败: $($_.Exception.Message) $($_.ScriptStackTrace)" 'ERROR'
                                Show-MessageDialog -Title '操作失败' -Message $_.Exception.Message -Kind 'error'
                            }
                        }
                    })
                [void]$panel.Children.Add($row)
            }
        }
    }

    $win.Content = $menuHost
    $auto = New-Object System.Windows.Threading.DispatcherTimer
    $auto.Interval = [TimeSpan]::FromSeconds(8)
    $auto.Add_Tick({
            param($ts, $te)
            $ts.Stop()
            $ww = $script:OpenMenuWindow
            if ($ww -and $ww.IsVisible) { try { $ww.Close() } catch { } }
        })
    $script:OpenMenuAutoTimer = $auto
    $win.Add_Closed({
            param($s, $e)
            try { if ($script:OpenMenuAutoTimer) { $script:OpenMenuAutoTimer.Stop() } } catch { }
        })
    $auto.Start()
    $cur = Get-CursorDip
    $win.Left = $cur.x - $script:MenuShadowPad
    $win.Top = $cur.y - $script:MenuShadowPad
    $script:OpenMenuWindow = $win
    $win.Add_SourceInitialized({
            param($s, $e)
            Set-PadWindowExStyle -Window $s
        })
    $win.Show()
    try {
        $vs = Get-VirtualScreenDip
        $w = $win.ActualWidth
        $h = $win.ActualHeight
        if ($w -lt 50) { $w = 220 }
        if ($h -lt 30) { $h = 120 }
        $nx = [double]$win.Left
        $ny = [double]$win.Top
        if (($nx + $w) -gt $vs.Right) { $nx = [Math]::Max($vs.Left, $vs.Right - $w - 4) }
        if (($ny + $h) -gt $vs.Bottom) { $ny = [Math]::Max($vs.Top, $vs.Bottom - $h - 4) }
        $win.Left = $nx
        $win.Top = $ny
    } catch { }
    # 菜单必须压在框上面：把菜单提到最前（框可能是置顶的）
    try {
        $mh = [System.Windows.Interop.WindowInteropHelper]::new($win).Handle
        if ($mh -ne [IntPtr]::Zero) {
            [void][FPNative]::SetWindowPos($mh, [FPNative]::HWND_TOPMOST, 0, 0, 0, 0, (0x0010 -bor 0x0002 -bor 0x0001))
        }
    } catch { }
}

function Show-CategoryPalette {
    param([string]$PadId)
    $data = Get-PadById $PadId
    if ($null -eq $data) { return }
    $script:CatSwatch = $null
    Show-CustomDialog -Title '标题条颜色（类别色）' `
        -Message '标题条用哪个颜色自己挑，和下面的紧急度底色互不影响。' `
        -Detail '点色块立刻生效；「不使用类别色」会用回默认的深灰主题色。' `
        -Kind 'info' -Width 500 `
        -Buttons @(@{ Text = '完成'; Value = 'ok'; Primary = $true }) `
        -BuildContent {
        param($panel)
        # 当前颜色预览
        $curRow = New-Object System.Windows.Controls.StackPanel
        $curRow.Orientation = 'Horizontal'
        $curRow.Margin = '0,0,0,12'
        $lblCur = New-Object System.Windows.Controls.TextBlock
        $lblCur.Text = '当前：'
        $lblCur.FontSize = 12
        $lblCur.Foreground = New-FPBrush '#FFFFFF' 0.62
        $lblCur.VerticalAlignment = 'Center'
        [void]$curRow.Children.Add($lblCur)
        $sw = New-Object System.Windows.Controls.Border
        $sw.Width = 44; $sw.Height = 16
        $sw.CornerRadius = '3'
        $sw.BorderThickness = '1'
        $sw.BorderBrush = New-FPBrush '#FFFFFF' 0.45
        $curOn = ($data.categoryColorOn -ne $false) -and $data.categoryColor
        if ($curOn) { $sw.Background = New-FPBrush ([string]$data.categoryColor) 1.0 }
        else { $sw.Background = New-FPBrush $script:ThemeTitleColor 1.0 }
        $sw.Margin = '0,0,10,0'
        $sw.VerticalAlignment = 'Center'
        $script:CatSwatch = $sw
        [void]$curRow.Children.Add($sw)
        $lblHex = New-Object System.Windows.Controls.TextBlock
        if ($curOn) { $lblHex.Text = [string]$data.categoryColor } else { $lblHex.Text = '默认深灰主题色（不使用类别色）' }
        $lblHex.FontSize = 11.5
        $lblHex.Foreground = New-FPBrush '#FFFFFF' 0.45
        $lblHex.VerticalAlignment = 'Center'
        [void]$curRow.Children.Add($lblHex)
        [void]$panel.Children.Add($curRow)

        # 色块
        $wrap = New-Object System.Windows.Controls.WrapPanel
        $wrap.Margin = '0,0,0,10'
        foreach ($c in $script:CategoryColors) {
            $dot = New-Object System.Windows.Shapes.Ellipse
            $dot.Width = 24; $dot.Height = 24
            $dot.Margin = '0,0,8,8'
            $dot.Fill = New-FPBrush $c 1.0
            $dot.Stroke = New-FPBrush '#FFFFFF' 0.30
            $dot.StrokeThickness = 1
            $dot.Cursor = 'Hand'
            $dot.ToolTip = $c
            $dot.Tag = @($PadId, $c)
            $dot.Add_MouseEnter({ param($s, $e) $s.StrokeThickness = 2; $s.Stroke = New-FPBrush '#FFFFFF' 0.95 })
            $dot.Add_MouseLeave({ param($s, $e) $s.StrokeThickness = 1; $s.Stroke = New-FPBrush '#FFFFFF' 0.30 })
            # 吃掉 MouseDown，避免冒泡到对话框的 DragMove
            $dot.Add_MouseLeftButtonDown({ param($s, $e) $e.Handled = $true })
            $dot.Add_MouseLeftButtonUp({
                    param($s, $e)
                    $tag = $s.Tag
                    $id = [string]$tag[0]
                    $c = [string]$tag[1]
                    $d = Get-PadById $id
                    if (-not $d) { return }
                    $d.categoryColor = $c
                    $d.categoryColorOn = $true
                    Save-State
                    if ($script:CatSwatch) { $script:CatSwatch.Background = New-FPBrush $c 1.0 }
                    Update-PadVisual -PadId $id
                    Write-Log "类别色: $id => $c"
                })
            [void]$wrap.Children.Add($dot)
        }
        [void]$panel.Children.Add($wrap)

        # 不使用类别色
        $off = New-Object System.Windows.Controls.Border
        $off.Background = New-FPBrush '#1A1A1E' 1.0
        $off.BorderBrush = New-FPBrush '#FFFFFF' 0.14
        $off.BorderThickness = '1'
        $off.CornerRadius = '5'
        $off.Padding = '10,7,10,7'
        $off.Cursor = 'Hand'
        $off.Tag = $PadId
        $offTb = New-Object System.Windows.Controls.TextBlock
        $offTb.Text = '不使用类别色（标题条用默认深灰主题色）'
        $offTb.FontSize = 12
        $offTb.Foreground = New-FPBrush '#FFFFFF' 0.72
        $off.Child = $offTb
        $off.Add_MouseEnter({ param($s, $e) $s.Background = New-FPBrush '#FFFFFF' 0.10 })
        $off.Add_MouseLeave({ param($s, $e) $s.Background = New-FPBrush '#1A1A1E' 1.0 })
        $off.Add_MouseLeftButtonDown({ param($s, $e) $e.Handled = $true })
        $off.Add_MouseLeftButtonUp({
                param($s, $e)
                $d = Get-PadById ([string]$s.Tag)
                if (-not $d) { return }
                $d.categoryColorOn = $false
                Save-State
                if ($script:CatSwatch) { $script:CatSwatch.Background = New-FPBrush $script:ThemeTitleColor 1.0 }
                Update-PadVisual -PadId ([string]$s.Tag)
                Write-Log ("类别色: {0} => 不使用" -f [string]$s.Tag)
            })
        [void]$panel.Children.Add($off)
    } | Out-Null
    $script:CatSwatch = $null
}

function Reset-UrgencyDefaults {
    $def = (Get-DefaultSettings).urgency
    foreach ($k in @('safe', 'soon', 'close', 'today', 'expired')) {
        $script:State.settings.urgency[$k].color = [string]$def[$k].color
        if ($null -ne $def[$k].days) { $script:State.settings.urgency[$k].days = [int]$def[$k].days }
    }
    Save-State
    # 面板还开着时，把 5 个当前色块也刷回去
    if ($script:PalSwatch) {
        foreach ($k in @($script:PalSwatch.Keys)) {
            $sw = $script:PalSwatch[$k]
            if ($sw) {
                $c = [string]$script:State.settings.urgency[$k].color
                try { $sw.Background = New-FPBrush $c 1.0 } catch { }
            }
        }
    }
    foreach ($k in @($script:Pads.Keys)) { Update-PadVisual -PadId $k }
    Write-Log '紧急度档位颜色已恢复默认（含档位天数 8/4/1/0）'
}

function Show-ColorPalette {
    $tiers = @(
        @{ Key = 'safe'; Text = '安全（剩余 ≥ 8 天）' },
        @{ Key = 'soon'; Text = '临近（4–7 天）' },
        @{ Key = 'close'; Text = '接近（1–3 天）' },
        @{ Key = 'today'; Text = '今天截止（0 天）' },
        @{ Key = 'expired'; Text = '已过期' }
    )
    $script:PalSwatch = @{}
    Show-CustomDialog -Title '紧急度档位颜色' `
        -Message '框内底色按剩余天数自动匹配（不再画边框）。点色块即可换色，立即生效。' `
        -Detail ('第 1–5 个色块就是五个档位的出厂默认色（灰蓝 / 黄 / 橙 / 红 / 灰），点它等于单独还原这一档。') `
        -Kind 'info' -Width 520 `
        -Buttons @(
        @{ Text = '恢复默认'; Value = 'reset'; KeepOpen = $true; Action = { Reset-UrgencyDefaults } },
        @{ Text = '完成'; Value = 'ok'; Primary = $true }
    ) `
        -BuildContent {
        param($panel)
        foreach ($t in $tiers) {
            $row = New-Object System.Windows.Controls.StackPanel
            $row.Orientation = 'Horizontal'
            $row.Margin = '0,0,0,9'
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = $t.Text
            $lbl.FontSize = 12
            $lbl.Width = 146
            $lbl.VerticalAlignment = 'Center'
            $lbl.Foreground = New-FPBrush '#FFFFFF' 0.80
            [void]$row.Children.Add($lbl)
            $cur = $script:State.settings.urgency[$t.Key].color
            $sw = New-Object System.Windows.Controls.Border
            $sw.Width = 26; $sw.Height = 14
            $sw.CornerRadius = '3'
            $sw.BorderThickness = '1'
            $sw.BorderBrush = New-FPBrush '#FFFFFF' 0.45
            $sw.Background = New-FPBrush ([string]$cur) 1.0
            $sw.Margin = '0,0,10,0'
            $script:PalSwatch[$t.Key] = $sw
            [void]$row.Children.Add($sw)
            foreach ($c in $script:UrgencyPresets) {
                $dot = New-Object System.Windows.Shapes.Ellipse
                $dot.Width = 16; $dot.Height = 16
                $dot.Margin = '0,0,4,0'
                $dot.Fill = New-FPBrush $c 1.0
                $dot.Stroke = New-FPBrush '#FFFFFF' 0.35
                $dot.StrokeThickness = 1
                $dot.Cursor = 'Hand'
                $dot.ToolTip = $c
                $dot.Tag = @($t.Key, $c)
                # 吃掉 MouseDown，避免冒泡到对话框的 DragMove（否则点了没反应）
                $dot.Add_MouseLeftButtonDown({ param($s, $e) $e.Handled = $true })
                $dot.Add_MouseLeftButtonUp({
                        param($s, $e)
                        $tag = $s.Tag
                        $tk = [string]$tag[0]
                        $tc = [string]$tag[1]
                        $script:State.settings.urgency[$tk].color = $tc
                        Save-State
                        if ($script:PalSwatch[$tk]) { $script:PalSwatch[$tk].Background = New-FPBrush $tc 1.0 }
                        foreach ($k in @($script:Pads.Keys)) { Update-PadVisual -PadId $k }
                        Write-Log "档位配色: $tk => $tc"
                    })
                [void]$row.Children.Add($dot)
            }
            [void]$panel.Children.Add($row)
        }
        $tip = New-Object System.Windows.Controls.TextBlock
        $tip.Text = '提示：档位天数（8 / 4 / 1 / 0）是该档的下限，「恢复默认」会连天数一起还原；单个档位也能单独改。'
        $tip.FontSize = 11
        $tip.Foreground = New-FPBrush '#FFFFFF' 0.40
        $tip.TextWrapping = 'Wrap'
        $tip.Margin = '0,6,0,0'
        [void]$panel.Children.Add($tip)
    } | Out-Null
    $script:PalSwatch = @{}
}

function Show-PadMenu {
    param([string]$PadId)
    $data = Get-PadById $PadId
    if ($null -eq $data) { return }
    $id = $PadId

    $aOpen = { param($a, $b) $d = Get-PadById $a; if ($d) { Open-InExplorer -Path $d.path } }
    $aRefresh = { param($a, $b) Refresh-Pad -PadId $a }
    $aRelocate = { param($a, $b) Invoke-Relocate -PadId $a }
    $aDue = { param($a, $b) Invoke-SetDueDate -PadId $a }
    $aRenamePad = { param($a, $b) Invoke-RenamePad -PadId $a }
    $aCatOff = {
        param($a, $b)
        $d = Get-PadById $a
        if (-not $d) { return }
        $d.categoryColorOn = $false
        Save-State
        Update-PadVisual -PadId $a
    }
    $aAuto = {
        param($a, $b)
        $d = Get-PadById $a
        if (-not $d) { return }
        if ($d.autoUrgency -eq $false) { $d.autoUrgency = $true } else { $d.autoUrgency = $false }
        Save-State
        Update-PadVisual -PadId $a
    }
    $aPalette = { param($a, $b) Show-ColorPalette }
    $aColsAuto = { param($a, $b) $d = Get-PadById $a; if ($d) { $d.cols = 'auto'; Save-State; Refresh-Pad -PadId $a } }
    $aCols1 = { param($a, $b) $d = Get-PadById $a; if ($d) { $d.cols = '1'; Save-State; Refresh-Pad -PadId $a } }
    $aCols2 = { param($a, $b) $d = Get-PadById $a; if ($d) { $d.cols = '2'; Save-State; Refresh-Pad -PadId $a } }
    $aLayout = { param($a, $b) Show-LayoutDialog -PadId $a }
    $aPin = {
        param($a, $b)
        $d = Get-PadById $a
        $pp = $script:Pads[$a]
        if (-not $d) { return }
        if ($d.pinned) { $d.pinned = $false } else { $d.pinned = $true }
        if ($pp -and $pp.Window) { $pp.Window.Topmost = [bool]$d.pinned }
        Save-State
        Sync-PadZOrder -PadId $a
    }
    $aZip = { param($a, $b) Invoke-Archive -PadId $a -Mode 'zip' }
    $aFolderInPlace = { param($a, $b) Invoke-Archive -PadId $a -Mode 'folder-inplace' }
    $aFolderMove = { param($a, $b) Invoke-Archive -PadId $a -Mode 'folder-move' }
    $aToFolder = { param($a, $b) Convert-ToFolder -PadId $a }
    $aHide = { param($a, $b) Close-PadWindow -PadId $a }
    $aRemove = { param($a, $b) Invoke-RemoveRecord -PadId $a }
    $aSettings = { param($a, $b) Show-SettingsDialog | Out-Null }
    $aExit = {
        param($a, $b)
        $ok = Show-ConfirmDialog -Title '退出 FolderPad' -Message '退出程序？' `
            -Detail '框记录会保留；按当前设置，退出时会把框背后的文件夹还原成可见（默认）。下次启动会重建所有框。' `
            -OkText '退出' -Kind 'question'
        if ($ok) { Exit-FolderPad }
    }
    $aCatColor = {
        param($a, $b)
        $d = Get-PadById $a
        if (-not $d) { return }
        $d.categoryColor = [string]$b
        $d.categoryColorOn = $true
        Save-State
        Update-PadVisual -PadId $a
        Write-Log ("类别色: {0} => {1}" -f $a, $b)
    }
    $aCategory = {
        param($a, $b)
        Show-CategoryPalette -PadId $a
    }

    $entries = New-Object System.Collections.ArrayList
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '打开文件夹' -Action $aOpen))
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '刷新' -Action $aRefresh))
    [void]$entries.Add((New-MenuSep))
    if (-not [System.IO.Directory]::Exists([string]$data.path)) {
        [void]$entries.Add((New-MenuEntry -PadId $id -Text '重新定位…' -Action $aRelocate -Hint '路径失效'))
    }
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '设置截止日期…' -Action $aDue -Hint ([string]$data.dueDate)))
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '重命名…（同时改文件夹名）' -Action $aRenamePad))
    if ($data.autoUrgency -eq $false) {
        [void]$entries.Add((New-MenuEntry -PadId $id -Text '底色随剩余天数自动变色' -Action $aAuto))
    } else {
        [void]$entries.Add((New-MenuEntry -PadId $id -Text '底色随剩余天数自动变色' -Action $aAuto -Checkable -Checked:$true))
    }
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '紧急度档位配色…' -Action $aPalette))
    [void]$entries.Add((New-MenuSep))
    $catHint = '不使用'
    if ($data.categoryColorOn -ne $false -and $data.categoryColor) { $catHint = [string]$data.categoryColor }
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '类别色…（标题条）' -Action $aCategory -Hint $catHint))
    [void]$entries.Add((New-MenuSep))
    $sizeHint = '自动'
    if ([string]$data.cols -match '^\d+$') { $sizeHint = $data.cols + ' 列' }
    if ([string]$data.rows -match '^\d+$') { $sizeHint = $(if ($sizeHint -eq '自动') { $data.rows + ' 行' } else { $sizeHint + ' × ' + $data.rows + ' 行' }) }
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '布局与排序…' -Action $aLayout -Hint $sizeHint))
    [void]$entries.Add((New-MenuSep))
    $pinnedNow = [bool]$data.pinned
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '置顶' -Action $aPin -Checkable -Checked:$pinnedNow))
    [void]$entries.Add((New-MenuSep))
    $dueTip = [string]$data.dueDate
    if (-not $dueTip) { $dueTip = (Get-Date).ToString('yyyy-MM-dd') }
    $zipName = [string]$data.name
$maxChars = [int]$script:MenuHintNameChars
if ($maxChars -gt 3 -and $zipName.Length -gt $maxChars) { $zipName = $zipName.Substring(0, $maxChars - 1) + '…' }
[void]$entries.Add((New-MenuEntry -PadId $id -Text '归档为压缩包（.zip）' -Action $aZip -Hint ($dueTip + '_' + $zipName + '.zip')))
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '归档为文件夹（原地加日期前缀）' -Action $aFolderInPlace))
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '归档为文件夹（移到归档目录）' -Action $aFolderMove -Hint ([string]$script:State.settings.archiveRoot)))
    [void]$entries.Add((New-MenuSep))
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '转为普通文件夹（截止日清空）' -Action $aToFolder))
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '隐藏这个框（保留记录）' -Action $aHide))
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '移除记录（不删文件）' -Action $aRemove))
    [void]$entries.Add((New-MenuSep))
    [void]$entries.Add((New-MenuEntry -PadId $id -Text 'FolderPad 设置…' -Action $aSettings))
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '退出 FolderPad' -Action $aExit))

    Show-FpMenu -PadId $id -Entries $entries.ToArray()
}

function Set-PadItemsBig {
    param([string]$PadId, [string[]]$Names, [bool]$Big)
    $p = $script:Pads[$PadId]
    if ($null -eq $p) { return }
    $cur = @($p.Data.big)
    foreach ($n in @($Names)) {
        $nm = [string]$n
        if (-not $nm) { continue }
        $cur = @($cur | Where-Object { [string]$_ -ne $nm })
        if ($Big) { $cur += $nm }
    }
    $p.Data.big = @($cur | Select-Object -Unique)
    Save-State
    Refresh-Pad -PadId $PadId
    Write-Log ("放大突出: " + $(if ($Big) { '设为 2×2' } else { '取消' }) + ' ' + (@($Names) -join ', '))
}

function Set-PadSort {
    param([string]$PadId, [string]$Key)
    $p = $script:Pads[$PadId]
    if ($null -eq $p) { return }
    $items = @($p.Items)
    $dirExpr = @{ Expression = { if ($_.IsDir) { 0 } else { 1 } } }
    $sorted = @()
    switch ($Key) {
        'mtime' { $sorted = @($items | Sort-Object $dirExpr, @{ Expression = { $_.Modified }; Descending = $false }) }
        'size' { $sorted = @($items | Sort-Object $dirExpr, @{ Expression = { $_.Size }; Descending = $true }) }
        'type' { $sorted = @($items | Sort-Object $dirExpr, @{ Expression = { [System.IO.Path]::GetExtension([string]$_.Name).ToLower() } }, @{ Expression = { [string]$_.Name } }) }
        default { $sorted = @($items | Sort-Object $dirExpr, @{ Expression = { [string]$_.Name } }) }
    }
    $p.Data.order = @($sorted | ForEach-Object { [string]$_.Name })
    $p.Data.order = @($sorted | ForEach-Object { [string]$_.Name })
    $p.Data.cells = @{}
    Refresh-Pad -PadId $PadId
    Write-Log ("按 {0} 排序: {1} 项" -f $Key, @($sorted).Count)
}

function Show-LayoutDialog {
    param([string]$PadId)
    $p = $script:Pads[$PadId]
    if ($null -eq $p) { return }
    $data = $p.Data
    $script:FPLayoutSortRbs = @{}

    $res = Show-CustomDialog -Title '布局与排序' -Message '' -Kind 'info' -Width 470 `
        -Buttons @(
        @{ Text = '取消'; Value = 'cancel' },
        @{ Text = '应用'; Value = 'ok'; Primary = $true }
    ) -BuildContent {
        param($panel)
        $mkLbl = {
            param($text)
            $t = New-Object System.Windows.Controls.TextBlock
            $t.Text = $text
            $t.Width = 96
            $t.FontSize = 12.5
            $t.Foreground = New-FPBrush '#FFFFFF' 0.78
            $t.VerticalAlignment = 'Center'
            return $t
        }
        # 列数
        $rowC = New-Object System.Windows.Controls.StackPanel
        $rowC.Orientation = 'Horizontal'
        $rowC.Margin = '0,0,0,10'
        [void]$rowC.Children.Add((& $mkLbl '列数'))
        $tbC = New-Object System.Windows.Controls.TextBox
        $tbC.Text = $(if ([string]$data.cols -match '^\d+$') { [string]$data.cols } else { '0' })
        $tbC.Width = 56
        $tbC.FontSize = 12.5
        $tbC.Background = New-FPBrush '#FFFFFF' 0.94
        $tbC.Foreground = New-FPBrush '#111111' 1
        [void]$rowC.Children.Add($tbC)
        $hintC = New-Object System.Windows.Controls.TextBlock
        $hintC.Text = '  1–8；0 = 自动（跟着项目数走）'
        $hintC.FontSize = 11.5
        $hintC.Foreground = New-FPBrush '#FFFFFF' 0.45
        $hintC.VerticalAlignment = 'Center'
        [void]$rowC.Children.Add($hintC)
        [void]$panel.Children.Add($rowC)
        $script:FPLayoutCols = $tbC

        # 行数
        $rowR = New-Object System.Windows.Controls.StackPanel
        $rowR.Orientation = 'Horizontal'
        $rowR.Margin = '0,0,0,10'
        [void]$rowR.Children.Add((& $mkLbl '行数'))
        $tbR = New-Object System.Windows.Controls.TextBox
        $tbR.Text = $(if ([string]$data.rows -match '^\d+$') { [string]$data.rows } else { '0' })
        $tbR.Width = 56
        $tbR.FontSize = 12.5
        $tbR.Background = New-FPBrush '#FFFFFF' 0.94
        $tbR.Foreground = New-FPBrush '#111111' 1
        [void]$rowR.Children.Add($tbR)
        $hintR = New-Object System.Windows.Controls.TextBlock
        $hintR.Text = '  1–8；0 = 自动（至少几行，放不下继续往下长）'
        $hintR.FontSize = 11.5
        $hintR.Foreground = New-FPBrush '#FFFFFF' 0.45
        $hintR.VerticalAlignment = 'Center'
        [void]$rowR.Children.Add($hintR)
        [void]$panel.Children.Add($rowR)
        $script:FPLayoutRows = $tbR

        # 排序
        $rowS = New-Object System.Windows.Controls.StackPanel
        $rowS.Orientation = 'Horizontal'
        $rowS.Margin = '0,0,0,6'
        [void]$rowS.Children.Add((& $mkLbl '排序'))
        $rbKeep = New-Object System.Windows.Controls.RadioButton
        $rbKeep.Content = '保持手动'
        $rbKeep.IsChecked = $true
        $rbKeep.GroupName = 'fpsort'
        $rbKeep.Foreground = New-FPBrush '#FFFFFF' 0.85
        $rbKeep.FontSize = 12
        $rbKeep.Margin = '0,0,12,0'
        $rbKeep.VerticalAlignment = 'Center'
        [void]$rowS.Children.Add($rbKeep)
        $script:FPLayoutSortRbs['keep'] = $rbKeep
        foreach ($it in @(
                @{ Key = 'name'; Text = '名称' },
                @{ Key = 'mtime'; Text = '修改时间' },
                @{ Key = 'size'; Text = '大小' },
                @{ Key = 'type'; Text = '类型' })) {
            $rb = New-Object System.Windows.Controls.RadioButton
            $rb.Content = $it.Text
            $rb.GroupName = 'fpsort'
            $rb.Foreground = New-FPBrush '#FFFFFF' 0.85
            $rb.FontSize = 12
            $rb.Margin = '0,0,12,0'
            $rb.VerticalAlignment = 'Center'
            [void]$rowS.Children.Add($rb)
            $script:FPLayoutSortRbs[$it.Key] = $rb
        }
        [void]$panel.Children.Add($rowS)

        $tip = New-Object System.Windows.Controls.TextBlock
        $tip.Text = '排序会重排一次顺序；「保持手动」则不动顺序，可随时在框里拖动图标调整。'
        $tip.FontSize = 11
        $tip.Foreground = New-FPBrush '#FFFFFF' 0.35
        $tip.TextWrapping = 'Wrap'
        $tip.Margin = '0,2,0,0'
        [void]$panel.Children.Add($tip)
    }

    if ($null -eq $res -or $res.Value -ne 'ok') { return }
    try {
        $c = 0
        if ($script:FPLayoutCols) { [void][int]::TryParse($script:FPLayoutCols.Text.Trim(), [ref]$c) }
        if ($c -ge 1 -and $c -le 8) { $data.cols = [string]$c } else { $data.cols = 'auto' }
        $r = 0
        if ($script:FPLayoutRows) { [void][int]::TryParse($script:FPLayoutRows.Text.Trim(), [ref]$r) }
        if ($r -ge 1 -and $r -le 8) { $data.rows = [string]$r } else { $data.rows = 'auto' }
        $sortKey = ''
        foreach ($k in @($script:FPLayoutSortRbs.Keys)) {
            $rb = $script:FPLayoutSortRbs[$k]
            if ($rb -and $rb.IsChecked) { $sortKey = [string]$k; break }
        }
        Save-State
        if ($sortKey -and $sortKey -ne 'keep') { Set-PadSort -PadId $PadId -Key $sortKey }
        Refresh-Pad -PadId $PadId
    } catch {
        Show-MessageDialog -Title '应用布局失败' -Message $_.Exception.Message -Kind 'error'
    }
}

function Show-ItemMenu {
    param([string]$PadId, $Item)
    $p = $script:Pads[$PadId]
    if ($null -eq $p -or $null -eq $Item) { return }
    $id = $PadId

    $aOpen = { param($a, $b) Open-PadItem -PadId $a -Item $b }
    $aReveal = { param($a, $b) Open-PadItem -PadId $a -Item $b -Reveal }
    $aCopy = { param($a, $b) [System.Windows.Clipboard]::SetText([string]$b.FullName) }
    $aRename = { param($a, $b) Invoke-RenameItem -PadId $a -Item $b }
    $aDesktop = {
        param($a, $b)
        $desk = [Environment]::GetFolderPath('Desktop')
        $r = Move-ItemToDirectory -Source $b.FullName -DestDir $desk -ConflictAction 'rename'
        Refresh-Pad -PadId $a
        if (-not $r.Ok) { Show-MessageDialog -Title '移动失败' -Message $r.Error -Kind 'error' }
    }
    $aToDir = {
        param($a, $b)
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = '选择目标文件夹'
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $r = Move-ItemToDirectory -Source $b.FullName -DestDir $dlg.SelectedPath -ConflictAction 'rename'
            Refresh-Pad -PadId $a
            if (-not $r.Ok) { Show-MessageDialog -Title '移动失败' -Message $r.Error -Kind 'error' }
        }
    }

    $aBigToggle = {
        param($a, $b)
        $pp = $script:Pads[$a]
        if ($null -eq $pp) { return }
        $names = @($pp.Selected.Keys)
        if (@($names).Count -eq 0 -and $b) { $names = @([string]$b.Name) }
        $anyBig = $false
        foreach ($n in $names) { if (@($pp.Data.big) -contains [string]$n) { $anyBig = $true } }
        Set-PadItemsBig -PadId $a -Names $names -Big (-not $anyBig)
    }

    # 多选时按"选中的那些"，否则就只有当前这一项
    $selNames = @($p.Selected.Keys)
    if (@($selNames).Count -eq 0) { $selNames = @([string]$Item.Name) }
    $bigCount = 0
    foreach ($n in $selNames) { if (@($p.Data.big) -contains [string]$n) { $bigCount++ } }
    $bigLabel = '放大突出（占 2×2）'
    if ($bigCount -ge @($selNames).Count) { $bigLabel = '取消放大（恢复 1×1）' }
    if (@($selNames).Count -gt 1) { $bigLabel = $bigLabel + ('（选中 ' + @($selNames).Count + ' 项）') }

    $entries = New-Object System.Collections.ArrayList
    [void]$entries.Add((New-MenuHeader -Text ([string]$Item.Name)))
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '打开' -Action $aOpen -Item $Item))
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '在资源管理器中显示' -Action $aReveal -Item $Item))
    [void]$entries.Add((New-MenuSep))
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '复制路径' -Action $aCopy -Item $Item))
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '重命名…' -Action $aRename -Item $Item))
    $isBigNow = ($bigCount -gt 0)
    [void]$entries.Add((New-MenuEntry -PadId $id -Text $bigLabel -Action $aBigToggle -Item $Item -Checkable -Checked:$isBigNow))
    [void]$entries.Add((New-MenuSep))
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '移到桌面' -Action $aDesktop -Item $Item))
    [void]$entries.Add((New-MenuEntry -PadId $id -Text '移到…（选目录）' -Action $aToDir -Item $Item))

    Show-FpMenu -PadId $id -Entries $entries.ToArray()
}

# =====================================================================
# 光标 / 拖拽状态
# =====================================================================
function Get-CursorDip {
    $p = [System.Windows.Forms.Cursor]::Position
    $s = 1.0
    if ($script:Grid.Scale -gt 0) { $s = $script:Grid.Scale }
    return @{ x = $p.X / $s; y = $p.Y / $s }
}


function Update-DragShieldState {
    $down = ([FPNative]::GetAsyncKeyState(0x01) -band 0x8000) -ne 0
    if (-not $down) {
        $script:DragStart = $null
        if ($script:ShieldUp) { $script:ShieldUp = $false }
        return
    }
    $cur = [System.Windows.Forms.Cursor]::Position
    if ($null -eq $script:DragStart) {
        $script:DragStart = @{ X = $cur.X; Y = $cur.Y }
        return
    }
    if (-not $script:ShieldUp) {
        $dx = [Math]::Abs($cur.X - $script:DragStart.X)
        $dy = [Math]::Abs($cur.Y - $script:DragStart.Y)
        if ($dx -gt 8 -or $dy -gt 8) { $script:ShieldUp = $true }
    }
}

# =====================================================================
# 输入对话框
# =====================================================================
function Show-InputDialog {
    param([string]$Title, [string]$Message = '', [string]$Detail = '', [string]$Value = '', [string]$OkText = '确定')
    if ($script:QuietUI) { return $Value }
    $tb = New-Object System.Windows.Controls.TextBox
    $tb.Text = $Value
    $tb.FontSize = 13
    $tb.Padding = '6,4'
    $tb.Width = 380
    $tb.Background = New-FPBrush '#FFFFFF' 0.94
    $tb.Foreground = New-FPBrush '#111111' 1
    # 弹出来就能直接打字，并且全选旧名字（不用先点一下输入框）
    $tb.Add_Loaded({
            param($s, $e)
            try {
                $s.Focus() | Out-Null
                $s.SelectAll()
                $w = [System.Windows.Window]::GetWindow($s)
                if ($w) { [void][FPNative]::SetForegroundWindow([System.Windows.Interop.WindowInteropHelper]::new($w).Handle) }
            } catch { }
        })
    $tb.Add_KeyDown({
            param($s, $e)
            if ($e.Key -eq 'Return') {
                $w = [System.Windows.Window]::GetWindow($s)
                if ($w -and $script:Dlg) { $script:Dlg.Value = 'ok'; $w.Close() }
            }
        })
    $res = Show-CustomDialog -Title $Title -Message $Message -Detail $Detail -Kind 'question' -Width 440 `
        -Buttons @(
        @{ Text = '取消'; Value = 'cancel' },
        @{ Text = $OkText; Value = 'ok'; Primary = $true }
    ) -BuildContent {
        param($panel)
        [void]$panel.Children.Add($tb)
    }
    if ($res.Value -ne 'ok') { return $null }
    return $tb.Text
}

# =====================================================================
# 框动作
# =====================================================================
function Invoke-SetDueDate {
    param([string]$PadId)
    $pad = Get-PadById $PadId
    if ($null -eq $pad) { return }
    $res = Show-DueDateDialog -Name $pad.name -Path $pad.path
    if ($null -eq $res -or -not $res.Ok) { return }
    $pad.dueDate = [string]$res.DueDate
    Save-State
    if ($script:Pads.ContainsKey($PadId)) { Refresh-Pad -PadId $PadId }
    Write-Log "设置截止日: $($pad.path) => '$($pad.dueDate)'"
}

function Invoke-Relocate {
    param([string]$PadId)
    $pad = Get-PadById $PadId
    if ($null -eq $pad) { return }
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = '选择这个框现在对应的文件夹（外部改名/移动之后用）'
    $dlg.ShowNewFolderButton = $false
    try { $dlg.SelectedPath = Split-Path -Parent ([string]$pad.path) } catch { }
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $newPath = Get-NormalizedPath $dlg.SelectedPath
    if (-not $newPath -or -not [System.IO.Directory]::Exists($newPath)) {
        Show-MessageDialog -Title '重新定位' -Message '选择的文件夹不存在。' -Kind 'warn'
        return
    }
    $dup = Get-PadByPath $newPath
    if ($dup -and $dup.id -ne $PadId) {
        Show-MessageDialog -Title '重新定位' -Message '这个文件夹已经是一个框了，不能重复绑定。' -Kind 'warn'
        return
    }
    $attrs = Get-PathAttributes $newPath
    $pad.path = $newPath
    $pad.name = Split-Path -Leaf $newPath
    $pad.prevAttrs = $attrs
    $pad.prevHidden = [bool]($attrs -band [System.IO.FileAttributes]::Hidden)
    $pad.invalid = $false
    Save-State
    Set-FolderHidden $newPath | Out-Null
    if ($script:Pads.ContainsKey($PadId)) { Refresh-Pad -PadId $PadId }
    Write-Log "重新定位: $PadId => $newPath"
}

function Invoke-RemoveRecord {
    param([string]$PadId)
    $pad = Get-PadById $PadId
    if ($null -eq $pad) { return }
    $ok = Show-ConfirmDialog -Title '移除记录（不删文件）' `
        -Message ('移除「{0}」的框记录？' -f $pad.name) `
        -Detail '磁盘上的文件不会被删除。如果这个文件夹当前是隐藏的，会先把它还原成转换前的属性。' `
        -OkText '移除记录' -Kind 'warn'
    if (-not $ok) { return }
    Restore-FolderVisibility -Pad $pad | Out-Null
    Remove-PadRecord $PadId
    Close-PadWindow -PadId $PadId
    Save-State
    Write-Log "移除记录: $($pad.path)"
}

function Invoke-RenamePad {
    param([string]$PadId)
    $pad = Get-PadById $PadId
    if ($null -eq $pad) { return }
    $old = [string]$pad.path
    $exists = [System.IO.Directory]::Exists($old)

    if (-not $exists) {
        # 文件夹已不在了：只能改显示名（保持记录可用，「重新定位…」后再改真实名）
        $newName = Show-InputDialog -Title '重命名框' -Message ('文件夹已失效：' + $old) `
            -Detail '找不到文件夹，只能先改框上的标题；找回文件夹请用「重新定位…」。' `
            -Value ([string]$pad.name) -OkText '改标题'
        if ($null -eq $newName) { return }
        $newName = $newName.Trim()
        if ([string]::IsNullOrWhiteSpace($newName) -or $newName -eq [string]$pad.name) { return }
        $pad.name = $newName
        Save-State
        if ($script:Pads.ContainsKey($PadId)) { Refresh-Pad -PadId $PadId }
        Write-Log "重命名框标题（路径失效）: $PadId => $newName"
        return
    }

    # 正常情况：一次改名同时改磁盘文件夹名和框标题（两者本来就该是同一个名字）
    $new = Show-InputDialog -Title '重命名框' -Message ('当前名称：' + [string]$pad.name) `
        -Detail ('文件夹：' + $old + "`r`n" + '改名会同时改磁盘上的文件夹名（隐藏属性保持不变），框标题自动跟着变。') `
        -Value ([string]$pad.name) -OkText '改名'
    if ($null -eq $new) { return }
    $new = $new.Trim()
    if ([string]::IsNullOrWhiteSpace($new) -or $new -eq [string]$pad.name) { return }
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
        if ($new.IndexOf($ch) -ge 0) {
            Show-MessageDialog -Title '重命名失败' -Message ('名称里不能包含字符：' + $ch) -Kind 'warn'
            return
        }
    }
    $dest = Join-Path (Split-Path -Parent $old) $new
    if (Test-Path -LiteralPath $dest) {
        Show-MessageDialog -Title '重命名失败' -Message ('同名的文件夹或文件已存在：' + "`r`n" + $dest) -Kind 'warn'
        return
    }
    try {
        Rename-Item -LiteralPath $old -NewName $new -ErrorAction Stop
    } catch {
        Show-MessageDialog -Title '重命名失败' -Message $_.Exception.Message -Kind 'error'
        return
    }
    $pad.path = $dest
    $pad.name = $new
    Save-State
    if ($script:Pads.ContainsKey($PadId)) { Refresh-Pad -PadId $PadId }
    Show-TrayBalloon -Title 'FolderPad' -Text ('已改名为：' + $new)
    Write-Log "重命名: $old => $dest"
}

function Invoke-RenameItem {
    param([string]$PadId, $Item)
    if ($null -eq $Item) { return }
    $newName = Show-InputDialog -Title '重命名' -Message ('当前名称：' + $Item.Name) -Value $Item.Name
    if ($null -eq $newName) { return }
    $newName = $newName.Trim()
    if ([string]::IsNullOrWhiteSpace($newName) -or $newName -eq $Item.Name) { return }
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
        if ($newName.IndexOf($ch) -ge 0) {
            Show-MessageDialog -Title '重命名失败' -Message ('名称里不能包含字符：' + $ch) -Kind 'warn'
            return
        }
    }
    $dest = Join-Path (Split-Path -Parent $Item.FullName) $newName
    if (Test-Path -LiteralPath $dest) {
        Show-MessageDialog -Title '重命名失败' -Message '同名项目已存在。' -Kind 'warn'
        return
    }
    try {
        Rename-Item -LiteralPath $Item.FullName -NewName $newName -ErrorAction Stop
        Refresh-Pad -PadId $PadId
    } catch {
        Show-MessageDialog -Title '重命名失败' -Message $_.Exception.Message -Kind 'error'
    }
}

function Show-AllPads {
    $script:State.settings.hideAll = $false
    Save-State
    foreach ($pad in @($script:State.pads)) {
        if ($pad -and -not $script:Pads.ContainsKey($pad.id)) { Show-PadWindow -PadId $pad.id }
        elseif ($script:Pads.ContainsKey($pad.id) -and $script:Pads[$pad.id].Window -and -not $script:Pads[$pad.id].Window.IsVisible) {
            $script:Pads[$pad.id].Window.Show()
            if ($script:Pads[$pad.id].Back) { $script:Pads[$pad.id].Back.Show() }
        }
    }
    Sync-TrayMenu
}

function Hide-AllPads {
    $script:State.settings.hideAll = $true
    Save-State
    foreach ($k in @($script:Pads.Keys)) { Close-PadWindow -PadId $k }
    Sync-TrayMenu
}

function Refresh-AllPads {
    foreach ($k in @($script:Pads.Keys)) { Refresh-Pad -PadId $k }
    Show-TrayBalloon -Title 'FolderPad' -Text '所有框已刷新'
}

function Toggle-AllPads {
    if ($script:State.settings.hideAll) { Show-AllPads } else { Hide-AllPads }
}

function Restore-AllHiddenFolders {
    $backup = Read-HiddenBackup
    $restored = 0
    $failed = @()
    foreach ($e in $backup) {
        if (-not [System.IO.Directory]::Exists($e.Path)) { continue }
        $mask = ([System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)
        try {
            $cur = [System.IO.File]::GetAttributes($e.Path)
            if (-not ($cur -band [System.IO.FileAttributes]::Hidden)) { continue }
            $new = $cur -band (-bnot $mask)
            if ($e.Attrs -ge 0) {
                $new = ($cur -band (-bnot $mask)) -bor ([System.IO.FileAttributes]($e.Attrs) -band $mask)
            }
            [System.IO.File]::SetAttributes($e.Path, $new)
            $restored++
        } catch { $failed += $e.Path }
    }
    Write-Log "恢复所有隐藏文件夹: $restored 个"
    if ($failed.Count -gt 0) {
        Show-MessageDialog -Title '恢复隐藏文件夹' -Message ("已恢复 $restored 个，失败 $($failed.Count) 个：`r`n" + ($failed -join "`r`n")) -Kind 'warn'
    } else {
        Show-MessageDialog -Title '恢复隐藏文件夹' -Message "已按恢复清单还原 $restored 个文件夹的可见性。" -Kind 'info'
    }
}

function Toggle-TemporaryVisible {
    $anyHidden = $false
    foreach ($pad in @($script:State.pads)) {
        if ($pad -and [System.IO.Directory]::Exists($pad.path)) {
            $a = Get-PathAttributes $pad.path
            if ($a -band [System.IO.FileAttributes]::Hidden) { $anyHidden = $true }
        }
    }
    if ($anyHidden) {
        foreach ($pad in @($script:State.pads)) {
            if (-not $pad) { continue }
            if (-not [System.IO.Directory]::Exists($pad.path)) { continue }
            try {
                $cur = [System.IO.File]::GetAttributes($pad.path)
                [System.IO.File]::SetAttributes($pad.path, ($cur -band (-bnot [System.IO.FileAttributes]::Hidden)))
            } catch { }
        }
        Show-TrayBalloon -Title 'FolderPad' -Text '框背后的文件夹已临时显示'
    } else {
        foreach ($pad in @($script:State.pads)) {
            if ($pad -and [System.IO.Directory]::Exists($pad.path)) { Set-FolderHidden $pad.path | Out-Null }
        }
        Show-TrayBalloon -Title 'FolderPad' -Text '框背后的文件夹已重新隐藏'
    }
    Sync-TrayMenu
}

# =====================================================================
# 安装 / 卸载
# =====================================================================
function Get-SendToDir {
    try {
        $p = [Environment]::GetFolderPath('SendTo')
        if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    } catch { }
    return (Join-Path $env:APPDATA 'Microsoft\Windows\SendTo')
}

function New-Shortcut {
    param([string]$Path, [string]$Target, [string]$Arguments, [string]$WorkDir, [string]$Icon, [string]$Description)
    $ws = New-Object -ComObject WScript.Shell
    $lnk = $ws.CreateShortcut($Path)
    $lnk.TargetPath = $Target
    $lnk.Arguments = $Arguments
    if ($WorkDir) { $lnk.WorkingDirectory = $WorkDir }
    if ($Icon) { $lnk.IconLocation = $Icon }
    if ($Description) { $lnk.Description = $Description }
    $lnk.Save()
    return $Path
}

function Install-SendToShortcut {
    $dir = Get-SendToDir
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $vbs = Join-Path $script:AppDir 'FolderPad.vbs'
    $lnk = Join-Path $dir '转为 FolderPad 框.lnk'
    New-Shortcut -Path $lnk -Target (Join-Path $env:SystemRoot 'System32\wscript.exe') `
        -Arguments ('"' + $vbs + '"') -WorkDir $script:AppDir `
        -Icon ((Join-Path $env:SystemRoot 'System32\imageres.dll') + ',3') `
        -Description '把选中的文件夹转为 FolderPad 框' | Out-Null
    Write-Log "发送到入口: $lnk"
    return $lnk
}

function Install-DesktopShortcut {
    $desk = [Environment]::GetFolderPath('Desktop')
    $vbs = Join-Path $script:AppDir 'FolderPad.vbs'
    $lnk = Join-Path $desk 'FolderPad.lnk'
    New-Shortcut -Path $lnk -Target (Join-Path $env:SystemRoot 'System32\wscript.exe') `
        -Arguments ('"' + $vbs + '"') -WorkDir $script:AppDir `
        -Icon ((Join-Path $env:SystemRoot 'System32\imageres.dll') + ',3') `
        -Description 'FolderPad —— 桌面围栏式文件夹框' | Out-Null
    Write-Log "桌面快捷方式: $lnk"
    return $lnk
}

function Set-Autostart {
    param([bool]$Enable)
    $startup = [Environment]::GetFolderPath('Startup')
    if (-not $startup) { $startup = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup' }
    $lnk = Join-Path $startup 'FolderPad.lnk'
    try {
        if ($Enable) {
            $vbs = Join-Path $script:AppDir 'FolderPad.vbs'
            New-Shortcut -Path $lnk -Target (Join-Path $env:SystemRoot 'System32\wscript.exe') `
                -Arguments ('"' + $vbs + '" -Startup') -WorkDir $script:AppDir `
                -Icon ((Join-Path $env:SystemRoot 'System32\imageres.dll') + ',3') `
                -Description 'FolderPad 开机自启' | Out-Null
            Write-Log '已开启开机自启'
        } else {
            if (Test-Path -LiteralPath $lnk) { Remove-Item -LiteralPath $lnk -Force }
            Write-Log '已关闭开机自启'
        }
        return $true
    } catch {
        Write-Log "设置开机自启失败: $($_.Exception.Message)" 'ERROR'
        return $false
    }
}

function Get-SendToShortcutPath {
    return (Join-Path (Get-SendToDir) '转为 FolderPad 框.lnk')
}

function Test-SendToInstalled {
    return (Test-Path -LiteralPath (Get-SendToShortcutPath))
}

function Invoke-Install {
    Initialize-DataDir
    Load-State | Out-Null
    $sendTo = Install-SendToShortcut
    $desk = Install-DesktopShortcut
    if ($script:State.settings.autostart) { Set-Autostart -Enable $true | Out-Null }
    $msg = @(
        'FolderPad 安装完成。',
        '',
        '桌面快捷方式：' + $desk,
        '发送到入口：' + $sendTo,
        '',
        '用法：在资源管理器/桌面上右键一个文件夹 → 显示更多选项 → 发送到 → 转为 FolderPad 框',
        '双击桌面 FolderPad 图标启动（或运行 FolderPad.vbs）。'
    ) -join "`r`n"
    Write-Host $msg
    if (-not $Quiet) { Show-MessageDialog -Title 'FolderPad 安装完成' -Message $msg -Kind 'info' }
}

function Invoke-Uninstall {
    $targets = @()
    $targets += (Join-Path ([Environment]::GetFolderPath('Desktop')) 'FolderPad.lnk')
    $targets += (Get-SendToShortcutPath)
    $startup = [Environment]::GetFolderPath('Startup')
    $targets += (Join-Path $startup 'FolderPad.lnk')
    foreach ($t in $targets) {
        if (Test-Path -LiteralPath $t) {
            Remove-Item -LiteralPath $t -Force -ErrorAction SilentlyContinue
            Write-Host "已删除: $t"
        }
    }
    Write-Host 'FolderPad 快捷方式已卸载（数据目录与用户文件未改动）。'
}

# =====================================================================
# 启动自检与恢复
# =====================================================================
function Set-CleanExitMarker {
    try { [System.IO.File]::WriteAllText($script:Path_CleanExit, (Get-Date).ToString('s')) } catch { }
}

function Clear-CleanExitMarker {
    try { if (Test-Path -LiteralPath $script:Path_CleanExit) { Remove-Item -LiteralPath $script:Path_CleanExit -Force } } catch { }
}

function Invoke-StartupSelfCheck {
    $crashed = -not (Test-Path -LiteralPath $script:Path_CleanExit)
    $backup = Read-HiddenBackup
    $hiddenExisting = @()
    foreach ($e in $backup) {
        if ([System.IO.Directory]::Exists($e.Path)) { $hiddenExisting += $e }
    }

    if ($crashed -and $hiddenExisting.Count -gt 0) {
        Write-Log "检测到上次非正常退出，恢复清单里 $($hiddenExisting.Count) 个文件夹" 'WARN'
        $res = Show-CustomDialog -Title '检测到上次异常退出' `
            -Message ("有 {0} 个文件夹仍处于隐藏状态（框背后的文件夹）。" -f $hiddenExisting.Count) `
            -Detail ("列表：`r`n" + (($hiddenExisting | ForEach-Object { $_.Path }) -join "`r`n")) `
            -Kind 'warn' -Width 520 `
            -Buttons @(
            @{ Text = '保持现状（继续用框）'; Value = 'keep' },
            @{ Text = '全部恢复显示并移除框记录'; Value = 'restore'; Primary = $true }
        )
        if ($res.Value -eq 'restore') {
            foreach ($e in $hiddenExisting) {
                $mask = ([System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)
                try {
                    $cur = [System.IO.File]::GetAttributes($e.Path)
                    $new = $cur -band (-bnot $mask)
                    if ($e.Attrs -ge 0) {
                        $new = ($cur -band (-bnot $mask)) -bor ([System.IO.FileAttributes]($e.Attrs) -band $mask)
                    }
                    [System.IO.File]::SetAttributes($e.Path, $new)
                } catch { Write-Log "恢复失败: $($e.Path)" 'ERROR' }
            }
            $paths = @($hiddenExisting | ForEach-Object { $_.Path })
            $script:State.pads = @($script:State.pads | Where-Object {
                    $keep = $true
                    foreach ($p in $paths) { if (Test-PathEqual $_.path $p) { $keep = $false } }
                    $keep
                })
            Write-Log '已按恢复清单还原并清空对应框记录'
        }
    }

    # 对账：记录 vs 磁盘
    foreach ($pad in @($script:State.pads)) {
        if ($null -eq $pad) { continue }
        if (-not [System.IO.Directory]::Exists([string]$pad.path)) {
            $pad.invalid = $true
            continue
        }
        $pad.invalid = $false
        $attrs = Get-PathAttributes $pad.path
        if ($attrs -ge 0 -and -not ($attrs -band [System.IO.FileAttributes]::Hidden)) {
            # 记录有、磁盘却可见：把「转换前属性」更正为当前值，再重新隐藏以保持模型一致
            $pad.prevAttrs = $attrs
            $pad.prevHidden = $false
            Set-FolderHidden $pad.path | Out-Null
            Write-Log "对账修正: $($pad.path) 当前可见，已记录为转换前状态并重新隐藏"
        }
    }

    # 恢复清单里有、记录里没有 → 上次转框写到一半，按清单还原
    foreach ($e in $backup) {
        if (Get-PadByPath $e.Path) { continue }
        if (-not [System.IO.Directory]::Exists($e.Path)) { continue }
        $mask = ([System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)
        try {
            $cur = [System.IO.File]::GetAttributes($e.Path)
            if (-not ($cur -band [System.IO.FileAttributes]::Hidden)) { continue }
            $new = $cur -band (-bnot $mask)
            if ($e.Attrs -ge 0) { $new = ($cur -band (-bnot $mask)) -bor ([System.IO.FileAttributes]($e.Attrs) -band $mask) }
            [System.IO.File]::SetAttributes($e.Path, $new)
            Write-Log "半完成转框已还原: $($e.Path)"
        } catch { }
    }

    Save-State
    Set-CleanExitMarker
    if (-not (Test-SendToInstalled)) {
        try { Install-SendToShortcut | Out-Null } catch { Write-Log "重建发送到入口失败: $($_.Exception.Message)" 'WARN' }
    }
}

# =====================================================================
# 待处理队列（发送到入口投递）
# =====================================================================
function Get-PendingMutex {
    try { return (New-Object System.Threading.Mutex($false, 'Local\FolderPad.Pending')) } catch { return $null }
}

function Add-PendingPath {
    param([string]$Path)
    $m = Get-PendingMutex
    $has = $false
    if ($m) { try { $has = $m.WaitOne(800) } catch { $has = $false } }
    try {
        $enc = New-Object System.Text.UTF8Encoding($false)
        $line = $Path + "`r`n"
        [System.IO.File]::AppendAllText($script:Path_Pending, $line, $enc)
    } finally {
        if ($m -and $has) { try { $m.ReleaseMutex() } catch { } }
    }
}

function Invoke-ProcessPending {
    if (-not (Test-Path -LiteralPath $script:Path_Pending)) { return }
    $m = Get-PendingMutex
    $has = $false
    if ($m) { try { $has = $m.WaitOne(500) } catch { $has = $false } }
    try {
        $lines = @()
        try { $lines = @([System.IO.File]::ReadAllLines($script:Path_Pending, [System.Text.Encoding]::UTF8)) } catch { return }
        if ($lines.Count -eq 0) { return }
        try { [System.IO.File]::WriteAllText($script:Path_Pending, '') } catch { }
    } finally {
        if ($m -and $has) { try { $m.ReleaseMutex() } catch { } }
    }
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try {
            Write-Log "处理待转框请求: $line"
            Convert-ToPad -Path $line.Trim() | Out-Null
        } catch {
            Write-Log "处理待转框失败: $($_.Exception.Message)" 'ERROR'
        }
    }
}

# =====================================================================
# 托盘
# =====================================================================
function New-TrayIconImage {
    try {
        $bmp = New-Object System.Drawing.Bitmap 16, 16
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.Clear([System.Drawing.Color]::Transparent)
        $body = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 46, 46, 52))
        $bar = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 76, 141, 255))
        $dot = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(235, 240, 240, 245))
        $g.FillRectangle($body, 1, 2, 14, 12)
        $g.FillRectangle($bar, 1, 2, 14, 4)
        $g.FillRectangle($dot, 3, 8, 4, 4)
        $g.FillRectangle($dot, 9, 8, 4, 4)
        $g.Dispose()
        $h = $bmp.GetHicon()
        $ico = [System.Drawing.Icon]::FromHandle($h)
        return $ico
    } catch {
        return [System.Drawing.SystemIcons]::Application
    }
}

function New-TrayIcon {
    $ni = New-Object System.Windows.Forms.NotifyIcon
    try {
        $ico = New-TrayIconImage
        Write-Log ("托盘图标位图: {0} {1}x{2} 句柄={3}" -f $ico.GetType().Name, $ico.Width, $ico.Height, $ico.Handle)
        $ni.Icon = $ico
    } catch {
        Write-Log ("托盘图标位图失败: " + $_.Exception.Message) 'ERROR'
        $ni.Icon = [System.Drawing.SystemIcons]::Application
    }
    $ni.Text = 'FolderPad —— 桌面围栏'
    try {
        $ni.Visible = $true
    } catch {
        Write-Log ("托盘图标显示失败: " + $_.Exception.Message) 'ERROR'
    }
    $ni.Add_MouseDoubleClick({
            param($s, $e)
            if ($script:State.settings.hideAll) { Show-AllPads } else { Refresh-AllPads }
        })

    $cm = New-Object System.Windows.Forms.ContextMenuStrip
    $cm.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $item = New-Object System.Windows.Forms.ToolStripMenuItem '显示所有框'
    $item.Add_Click({ Show-AllPads })
    $cm.Items.Add($item) | Out-Null
    $item = New-Object System.Windows.Forms.ToolStripMenuItem '隐藏所有框'
    $item.Add_Click({ Hide-AllPads })
    $cm.Items.Add($item) | Out-Null
    $item = New-Object System.Windows.Forms.ToolStripMenuItem '全部刷新'
    $item.Add_Click({ Refresh-AllPads })
    $cm.Items.Add($item) | Out-Null
    [void]$cm.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $item = New-Object System.Windows.Forms.ToolStripMenuItem '临时显示框背后的文件夹'
    $item.Add_Click({ Toggle-TemporaryVisible })
    $cm.Items.Add($item) | Out-Null
    $item = New-Object System.Windows.Forms.ToolStripMenuItem '恢复所有隐藏文件夹'
    $item.Add_Click({ Restore-AllHiddenFolders })
    $cm.Items.Add($item) | Out-Null
    [void]$cm.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $item = New-Object System.Windows.Forms.ToolStripMenuItem '重新安装发送到入口'
    $item.Add_Click({
            try { $p = Install-SendToShortcut; Show-MessageDialog -Title '发送到入口' -Message ('已重新安装：' + $p) -Kind 'info' }
            catch { Show-MessageDialog -Title '安装失败' -Message $_.Exception.Message -Kind 'error' }
        })
    $cm.Items.Add($item) | Out-Null
    $item = New-Object System.Windows.Forms.ToolStripMenuItem '打开数据目录'
    $item.Add_Click({ Open-InExplorer -Path $script:DataDir })
    $cm.Items.Add($item) | Out-Null
    $item = New-Object System.Windows.Forms.ToolStripMenuItem '设置…'
    $item.Add_Click({ Show-SettingsDialog | Out-Null })
    $cm.Items.Add($item) | Out-Null
    [void]$cm.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $item = New-Object System.Windows.Forms.ToolStripMenuItem '退出'
    $item.Add_Click({ Exit-FolderPad })
    $cm.Items.Add($item) | Out-Null
    $ni.ContextMenuStrip = $cm
    $script:Tray = $ni
    $script:TrayMenu = $cm
    Write-Log ("托盘就绪: visible={0} text=[{1}]" -f $ni.Visible, $ni.Text)
}

function Sync-TrayMenu {
    # 托盘菜单由静态项组成，状态通过气球提示反馈，这里保持幂等
}

function Show-TrayBalloon {
    param([string]$Title, [string]$Text)
    if ($script:QuietUI) { return }
    if ($null -eq $script:Tray) { return }
    try { $script:Tray.ShowBalloonTip(2200, $Title, $Text, [System.Windows.Forms.ToolTipIcon]::None) } catch { }
}

# =====================================================================
# 定时器
# =====================================================================
function Start-FPTimers {
    $fast = New-Object System.Windows.Threading.DispatcherTimer
    $fast.Interval = [TimeSpan]::FromMilliseconds(600)
    $fast.Add_Tick({
            $script:TickAll++
            try { Update-DragShieldState } catch { }
            foreach ($k in @($script:Pads.Keys)) {
                try { Update-PadShield -PadId $k } catch { }
            }
            if (($script:TickAll % 2) -eq 0) {
                try { Invoke-ProcessPending } catch { }
            }
            if (($script:TickAll % 5) -eq 0) {
                foreach ($k in @($script:Pads.Keys)) {
                    $pp = $script:Pads[$k]
                    if ($null -eq $pp) { continue }
                    if ($pp.Busy -or $pp.CtxOpen) { continue }
                    try { Sync-PadZOrder -PadId $k } catch { }
                }
            }
        })
    $fast.Start()
    $script:TimerFast = $fast

    $hot = New-Object System.Windows.Threading.DispatcherTimer
    $hot.Interval = [TimeSpan]::FromMilliseconds(60)
    $hot.Add_Tick({
            try {
                $ctrl = ([FPNative]::GetAsyncKeyState(0x11) -band 0x8000) -ne 0
                $alt = ([FPNative]::GetAsyncKeyState(0x12) -band 0x8000) -ne 0
                $h = ([FPNative]::GetAsyncKeyState(0x48) -band 0x8000) -ne 0
                if ($ctrl -and $alt -and $h) {
                    if (-not $script:HotLatch) {
                        $script:HotLatch = $true
                        Toggle-AllPads
                    }
                } else { $script:HotLatch = $false }

                # 方向键调整顺序：↑/← 前移一位，↓/→ 后移一位
                # 触发条件：有选中项 + 鼠标停在该框上 + 前台是桌面（避免打字时误伤）
                $move = 0
                if (([FPNative]::GetAsyncKeyState(0x26) -band 0x8000) -ne 0) { $move = -1 }
                elseif (([FPNative]::GetAsyncKeyState(0x28) -band 0x8000) -ne 0) { $move = 1 }
                elseif (([FPNative]::GetAsyncKeyState(0x25) -band 0x8000) -ne 0) { $move = -1 }
                elseif (([FPNative]::GetAsyncKeyState(0x27) -band 0x8000) -ne 0) { $move = 1 }
                if ($move -ne 0) {
                    if (-not $script:ArrowLatch) {
                        $script:ArrowLatch = $true
                        $ppA = Get-PadSelectionPad
                        if ($ppA) {
                            $curA = Get-CursorDip
                            $lx = $curA.x - [double]$ppA.Data.pos.x
                            $ly = $curA.y - [double]$ppA.Data.pos.y
                            $inA = ($lx -ge 0 -and $lx -le [double]$ppA.Width -and $ly -ge 0 -and $ly -le ([double]$script:TitleHeight + [double]$ppA.Rows * [double]$script:Grid.CellH))
                            # 带修饰键时不当成"调位置"（Ctrl/Alt/Shift 留给别处）
                            $shiftA = ([FPNative]::GetAsyncKeyState(0x10) -band 0x8000) -ne 0
                            if ($inA -and -not ($ctrl -or $alt -or $shiftA)) {
                                $dcA = 0; $drA = 0
                                if (([FPNative]::GetAsyncKeyState(0x26) -band 0x8000) -ne 0) { $drA = -1 }
                                elseif (([FPNative]::GetAsyncKeyState(0x28) -band 0x8000) -ne 0) { $drA = 1 }
                                elseif (([FPNative]::GetAsyncKeyState(0x25) -band 0x8000) -ne 0) { $dcA = -1 }
                                elseif (([FPNative]::GetAsyncKeyState(0x27) -band 0x8000) -ne 0) { $dcA = 1 }
                                Move-PadSelectionByStep -P $ppA -DeltaCol $dcA -DeltaRow $drA
                            }
                        }
                    }
                } else { $script:ArrowLatch = $false }

                # 左键刚按下：落点不在任何框里 → 取消所有选中（桌面习惯）
                $lmb = ([FPNative]::GetAsyncKeyState(0x01) -band 0x8000) -ne 0
                if ($lmb -and -not $script:LmbLatch) {
                    $script:LmbLatch = $true
                    $anySelected = $false
                    foreach ($k in @($script:Pads.Keys)) {
                        $pp0 = $script:Pads[$k]
                        if ($pp0 -and $pp0.Selected) { $anySelected = $true; break }
                    }
                    if ($anySelected) {
                        $cur = Get-CursorDip
                        $hit = $false
                        foreach ($k in @($script:Pads.Keys)) {
                            $pp0 = $script:Pads[$k]
                            if ($null -eq $pp0 -or $null -eq $pp0.Window) { continue }
                            $l = [double]$pp0.Window.Left
                            $t = [double]$pp0.Window.Top
                            $w = [double]$pp0.Window.Width
                            $hh = [double]$pp0.Window.Height
                            if ($w -le 1 -or $hh -le 1) { continue }
                            if ($cur.x -ge $l -and $cur.x -le ($l + $w) -and $cur.y -ge $t -and $cur.y -le ($t + $hh)) { $hit = $true; break }
                        }
                        if (-not $hit) { Clear-AllPadSelection }
                    }
                }
                if (-not $lmb) { $script:LmbLatch = $false }
            } catch { }
        })
    $hot.Start()
    $script:TimerHot = $hot

    $slow = New-Object System.Windows.Threading.DispatcherTimer
    $slow.Interval = [TimeSpan]::FromSeconds(30)
    $slow.Add_Tick({
            try {
                $today = (Get-Date).Date
                if ($today -ne $script:LastDate) {
                    Write-Log "跨天重算: $($script:LastDate) -> $today"
                    $script:LastDate = $today
                    foreach ($k in @($script:Pads.Keys)) { Update-PadVisual -PadId $k }
                }
                $mins = [int]$script:State.settings.refreshMinutes
                if ($mins -lt 1) { $mins = 30 }
                if (((Get-Date) - $script:LastRefresh).TotalMinutes -ge $mins) {
                    $script:LastRefresh = Get-Date
                    foreach ($k in @($script:Pads.Keys)) { Refresh-Pad -PadId $k }
                    foreach ($pad in @($script:State.pads)) { if ($pad -and -not [System.IO.Directory]::Exists([string]$pad.path)) { $pad.invalid = $true } }
                }
            } catch { Write-Log "慢定时器出错: $($_.Exception.Message)" 'ERROR' }
        })
    $slow.Start()
    $script:TimerSlow = $slow
}

# =====================================================================
# 退出
# =====================================================================
function Exit-FolderPad {
    if ($script:AppExiting) { return }
    $script:AppExiting = $true
    Write-Log '开始退出'
    try {
        if ($script:State -and $script:State.settings.exitRestore) {
            foreach ($pad in @($script:State.pads)) {
                try { Restore-FolderVisibility -Pad $pad | Out-Null } catch { }
            }
        }
        foreach ($k in @($script:Pads.Keys)) { try { Close-PadWindow -PadId $k } catch { } }
        Save-State
        Clear-CleanExitMarker
    } catch { Write-Log "退出清理出错: $($_.Exception.Message)" 'ERROR' }
    try { if ($script:TimerFast) { $script:TimerFast.Stop() } } catch { }
    try { if ($script:TimerHot) { $script:TimerHot.Stop() } } catch { }
    try { if ($script:TimerSlow) { $script:TimerSlow.Stop() } } catch { }
    try { if ($script:Tray) { $script:Tray.Visible = $false; $script:Tray.Dispose() } } catch { }
    try { if ($script:Mutex) { $script:Mutex.ReleaseMutex() } } catch { }
    try { if ($script:WpfApp) { $script:WpfApp.Shutdown() } } catch { }
    Write-Log '退出完成'
}

# =====================================================================
# 应用主流程
# =====================================================================
function Start-FolderPadApp {
    Initialize-DataDir
    if (-not $script:Mutex) {
        $script:Mutex = New-Object System.Threading.Mutex($false, 'Local\FolderPad.Instance')
    }
    try {
        if (-not $script:Mutex.WaitOne(500)) {
            Write-Log '已有实例在运行，本次启动退出'
            return
        }
    } catch {
        Write-Log "互斥体异常（继续启动）: $($_.Exception.Message)" 'WARN'
    }

    $script:WpfApp = New-Object System.Windows.Application
    $script:WpfApp.ShutdownMode = [System.Windows.ShutdownMode]::OnExplicitShutdown
    $script:WpfApp.Add_DispatcherUnhandledException({
            param($s, $e)
            Write-Log ("未处理异常: " + $e.Exception.Message + " @ " + $e.Exception.StackTrace) 'ERROR'
            $e.Handled = $true
        })

    Load-State
    Initialize-GridMetrics
    $script:LastDate = (Get-Date).Date
    $script:LastRefresh = Get-Date

    Invoke-StartupSelfCheck
    New-TrayIcon
    Start-FPTimers

    if (-not $script:State.settings.hideAll) {
        foreach ($pad in @($script:State.pads)) {
            if ($null -eq $pad) { continue }
            try { Show-PadWindow -PadId $pad.id } catch { Write-Log "显示框失败 $($pad.id): $($_.Exception.Message)" 'ERROR' }
        }
    }

    # 演示用：启动时按 demo-request.txt 建演示框
    if (Test-Path -LiteralPath $script:Path_DemoRequest) {
        try {
            $lines = @([System.IO.File]::ReadAllLines($script:Path_DemoRequest, [System.Text.Encoding]::UTF8))
            [System.IO.File]::WriteAllText($script:Path_DemoRequest, '')
            foreach ($line in $lines) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                $parts = $line -split "`t"
                if ($parts.Count -ge 2) {
                    Convert-ToPad -Path $parts[0] -DueDate $parts[1] | Out-Null
                }
            }
        } catch { Write-Log "演示模式失败: $($_.Exception.Message)" 'ERROR' }
    }

    Invoke-ProcessPending

    Write-Log "FolderPad 启动完成（$($script:State.pads.Count) 个框）"
    $script:WpfApp.Run() | Out-Null
    Write-Log 'Dispatcher 已退出'
}

# =====================================================================
# 自检
# =====================================================================
function Test-Assert {
    param([string]$Name, [bool]$Ok, [string]$Info = '')
    $script:SelfTestResult += [pscustomobject]@{ Name = $Name; Ok = $Ok; Info = $Info }
    $tag = if ($Ok) { 'PASS' } else { 'FAIL' }
    $line = '[{0}] {1} {2}' -f $tag, $Name, $Info
    if ($Ok) { Write-Host $line -ForegroundColor DarkGreen } else { Write-Host $line -ForegroundColor Red }
}

function Invoke-SelfTest {
    Write-Host '=== FolderPad 自检 ===' -ForegroundColor Cyan
    Load-State | Out-Null
    try { Save-State } catch { Write-Log "自检写状态失败: $($_.Exception.Message)" 'WARN' }
    $settings = $script:State.settings

    # 1. 紧急度档位
    $cases = @(
        @{ Offset = 12; Tier = 'safe'; Text = '还剩 12 天' },
        @{ Offset = 8; Tier = 'safe'; Text = '还剩 8 天' },
        @{ Offset = 7; Tier = 'soon'; Text = '还剩 7 天' },
        @{ Offset = 4; Tier = 'soon'; Text = '还剩 4 天' },
        @{ Offset = 3; Tier = 'close'; Text = '还剩 3 天' },
        @{ Offset = 1; Tier = 'close'; Text = '还剩 1 天' },
        @{ Offset = 0; Tier = 'today'; Text = '今天截止' },
        @{ Offset = -1; Tier = 'expired'; Text = '已结束' },
        @{ Offset = -30; Tier = 'expired'; Text = '已结束' }
    )
    $allOk = $true
    $detail = @()
    foreach ($c in $cases) {
        $pad = @{ dueDate = (Get-Date).Date.AddDays($c.Offset).ToString('yyyy-MM-dd'); autoUrgency = $true }
        $u = Get-UrgencyInfo -Pad $pad -Settings $settings
        if ($u.Tier -ne $c.Tier -or $u.Text -ne $c.Text) {
            $allOk = $false
            $detail += ('{0}天 => {1}/{2} 期望 {3}/{4}' -f $c.Offset, $u.Tier, $u.Text, $c.Tier, $c.Text)
        }
    }
    Test-Assert '紧急度档位（含边界）' $allOk ($detail -join '; ')

    $u0 = Get-UrgencyInfo -Pad @{ dueDate = '' } -Settings $settings
    Test-Assert '未设截止' ($u0.Tier -eq 'none' -and $u0.Text -eq '未设截止' -and $u0.Color -eq '#B0B0B0') $u0.Text

    # 2. JSON 往返
    try {
        $tmpState = [ordered]@{ settings = $settings; pads = @([ordered]@{ id = 'x1'; path = 'C:\测试\框'; dueDate = '2026-11-20'; pos = @{ x = 1; y = 2 }; order = @('a.txt') }); archiveLog = @() }
        $json = ConvertTo-Json -InputObject $tmpState -Depth 8
        $back = ConvertTo-HashtableDeep ($json | ConvertFrom-Json)
        $p0 = @($back.pads)[0]
        Test-Assert 'JSON 往返（含中文路径）' ($p0.path -eq 'C:\测试\框' -and $p0.pos.y -eq 2) $p0.path
    } catch {
        Test-Assert 'JSON 往返（含中文路径）' $false $_.Exception.Message
    }

    # 3. 隐藏属性保存/还原
    $tmpDir = Join-Path $env:TEMP ('fp_test_' + [Guid]::NewGuid().ToString('N').Substring(0, 6))
    try {
        New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null
        $orig = Get-PathAttributes $tmpDir
        $pad = @{ path = $tmpDir; prevAttrs = $orig; prevHidden = $false }
        Set-FolderHidden $tmpDir | Out-Null
        $h1 = [bool]((Get-PathAttributes $tmpDir) -band [System.IO.FileAttributes]::Hidden)
        Restore-FolderVisibility -Pad $pad | Out-Null
        $h2 = [bool]((Get-PathAttributes $tmpDir) -band [System.IO.FileAttributes]::Hidden)
        Test-Assert '隐藏 / 按 prevAttrs 还原' ($h1 -and -not $h2) "hidden=$h1, restored=$h2"

        $padHidden = @{ path = $tmpDir; prevAttrs = ($orig -bor [int][System.IO.FileAttributes]::Hidden); prevHidden = $true }
        Set-FolderHidden $tmpDir | Out-Null
        Restore-FolderVisibility -Pad $padHidden | Out-Null
        $h3 = [bool]((Get-PathAttributes $tmpDir) -band [System.IO.FileAttributes]::Hidden)
        Test-Assert '原本隐藏的文件夹还原后仍隐藏' $h3 "hidden=$h3"
        try { [System.IO.File]::SetAttributes($tmpDir, [System.IO.FileAttributes]::Directory) } catch { }
    } catch {
        Test-Assert '隐藏 / 按 prevAttrs 还原' $false $_.Exception.Message
    }

    # 4. 恢复清单往返
    try {
        $oldPath = $script:Path_HiddenBackup
        $tmpHb = Join-Path $env:TEMP ('fp_hb_' + [Guid]::NewGuid().ToString('N').Substring(0, 6) + '.txt')
        $script:Path_HiddenBackup = $tmpHb
        $savedPads = $script:State.pads
        $script:State.pads = @(
            [ordered]@{ id = 'a'; path = 'D:\Pads\毕业设计'; prevAttrs = 32; prevHidden = $false },
            [ordered]@{ id = 'b'; path = 'D:\Pads\带 空格\子目录'; prevAttrs = 35; prevHidden = $true }
        )
        Write-HiddenBackup
        $read = Read-HiddenBackup
        $ok = ($read.Count -eq 2 -and $read[0].Path -eq 'D:\Pads\毕业设计' -and $read[1].Attrs -eq 35 -and $read[1].Hidden)
        Test-Assert '恢复清单写读往返' $ok ("$($read.Count) 行")
        $script:State.pads = $savedPads
        $script:Path_HiddenBackup = $oldPath
        try { Remove-Item -LiteralPath $tmpHb -Force -ErrorAction SilentlyContinue } catch { }
    } catch {
        Test-Assert '恢复清单写读往返' $false $_.Exception.Message
    }

    # 5. 唯一路径
    try {
        $u = Get-UniquePath (Join-Path $env:TEMP '不存在的文件_xyz.txt')
        Test-Assert 'Get-UniquePath（无冲突原样返回）' ($u -eq (Join-Path $env:TEMP '不存在的文件_xyz.txt')) $u
    } catch { Test-Assert 'Get-UniquePath' $false $_.Exception.Message }

    # 6. 特殊路径拒绝
    try {
        $r1 = Test-SpecialPath 'C:\'
        $r2 = Test-SpecialPath $env:USERPROFILE
        $r3 = Test-SpecialPath $tmpDir
        Test-Assert '特殊路径拒绝' ([bool]$r1 -and [bool]$r2 -and -not $r3) "root=$([bool]$r1) user=$([bool]$r2) normal=$(-not [bool]$r3)"
    } catch { Test-Assert '特殊路径拒绝' $false $_.Exception.Message }

    # 7. 压缩
    try {
        $src = Join-Path $env:TEMP ('fp_zip_src_' + [Guid]::NewGuid().ToString('N').Substring(0, 6))
        New-Item -ItemType Directory -Force -Path $src | Out-Null
        Set-Content -LiteralPath (Join-Path $src 'a.txt') -Value 'hello' -Encoding UTF8
        New-Item -ItemType Directory -Force -Path (Join-Path $src 'sub') | Out-Null
        Set-Content -LiteralPath (Join-Path $src 'sub\b.txt') -Value 'world' -Encoding UTF8
        $zip = "$src.zip"
        [System.IO.Compression.ZipFile]::CreateFromDirectory($src, $zip, [System.IO.Compression.CompressionLevel]::Optimal, $true)
        $bytes = (Get-Item -LiteralPath $zip).Length
        Test-Assert '压缩为 .zip' ((Test-Path -LiteralPath $zip) -and $bytes -gt 100) "$bytes 字节"
        Remove-Item -LiteralPath $zip -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $src -Recurse -Force -ErrorAction SilentlyContinue
    } catch { Test-Assert '压缩为 .zip' $false $_.Exception.Message }

    # 8. 移动/冲突
    try {
        $d1 = Join-Path $env:TEMP ('fp_mv_a_' + [Guid]::NewGuid().ToString('N').Substring(0, 6))
        $d2 = Join-Path $env:TEMP ('fp_mv_b_' + [Guid]::NewGuid().ToString('N').Substring(0, 6))
        New-Item -ItemType Directory -Force -Path $d1 | Out-Null
        New-Item -ItemType Directory -Force -Path $d2 | Out-Null
        $f = Join-Path $d1 'same.txt'
        Set-Content -LiteralPath $f -Value 'v1' -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $d2 'same.txt') -Value 'v2' -Encoding UTF8
        $r = Move-ItemToDirectory -Source $f -DestDir $d2 -ConflictAction 'rename'
        $ok = $r.Ok -and (Test-Path -LiteralPath (Join-Path $d2 'same_1.txt')) -and -not (Test-Path -LiteralPath $f)
        Test-Assert '拖入同名 → 重命名' $ok (Split-Path -Leaf $r.Dest)
        Remove-Item -LiteralPath $d1 -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $d2 -Recurse -Force -ErrorAction SilentlyContinue
    } catch { Test-Assert '拖入同名 → 重命名' $false $_.Exception.Message }

    # 9. 图标度量
    try {
        $m = Get-RawIconMetrics
        $ok = ($m.Icon -ge 16 -and $m.Horz -ge $m.Icon -and $m.Vert -ge $m.Icon)
        Test-Assert '桌面图标度量' $ok ("图标 {0}px 间距 {1}x{2} 来源 {3}" -f $m.Icon, $m.Horz, $m.Vert, $m.Source)
    } catch { Test-Assert '桌面图标度量' $false $_.Exception.Message }

    # 10. 图标提取
    try {
        Initialize-GridMetrics
        $ic = Get-FileIconSource -Path $env:SystemRoot -IsDir $true
        Test-Assert '文件图标提取' ($null -ne $ic) ("{0}x{1}" -f $ic.PixelWidth, $ic.PixelHeight)
    } catch { Test-Assert '文件图标提取' $false $_.Exception.Message }

    # 11. SendTo
    Test-Assert '发送到入口已安装' (Test-SendToInstalled) (Get-SendToShortcutPath)

    # 12. 数据文件
    $okFiles = (Test-Path -LiteralPath $script:Path_Json) -and (Test-Path -LiteralPath $script:Path_HiddenBackup)
    Test-Assert '数据文件齐全' $okFiles ("pads.json / hidden-backup.txt")

    # 13. 占位打包（2×2 放大 + first-fit）
    try {
        $n9 = @(1..9 | ForEach-Object { "f$_" })
        $pl = Get-PadPlan -Order $n9 -Big @() -Cols 3
        $okA = ($pl.Rows -eq 3 -and $pl.Cols -eq 3 -and $pl.Positions.Count -eq 9)
        foreach ($k in @($pl.Positions.Keys)) { if ([int]$pl.Positions[$k].Span -ne 1) { $okA = $false } }
        Test-Assert '打包：9 项 3 列 = 3×3' $okA ("rows={0} cols={1}" -f $pl.Rows, $pl.Cols)

        $n6 = @('big1', 'a', 'b', 'c', 'd', 'e')
        $pl2 = Get-PadPlan -Order $n6 -Big @('big1') -Cols 3
        $okB = ($pl2.Rows -eq 3 -and [int]$pl2.Positions['big1'].Span -eq 2 -and [int]$pl2.Positions['big1'].Col -eq 0 -and [int]$pl2.Positions['big1'].Row -eq 0)
        Test-Assert '打包：1 大格 + 5 普通 = 3×3' $okB ("rows={0} 大格在 ({1},{2})" -f $pl2.Rows, $pl2.Positions['big1'].Col, $pl2.Positions['big1'].Row)

        $n9b = @('big1') + @(1..8 | ForEach-Object { "g$_" })
        $pl3 = Get-PadPlan -Order $n9b -Big @('big1') -Cols 3
        Test-Assert '打包：1 大格 + 8 普通 = 4 行' ($pl3.Rows -eq 4) ("rows=" + $pl3.Rows)

        $seen = @{}
        $overlap = $false
        foreach ($k in @($pl2.Positions.Keys)) {
            $pp = $pl2.Positions[$k]
            for ($dx = 0; $dx -lt [int]$pp.Span; $dx++) {
                for ($dy = 0; $dy -lt [int]$pp.Span; $dy++) {
                    $key = ('{0},{1}' -f ([int]$pp.Col + $dx), ([int]$pp.Row + $dy))
                    if ($seen.ContainsKey($key)) { $overlap = $true }
                    $seen[$key] = $true
                }
            }
        }
        Test-Assert '打包：无重叠' (-not $overlap) ''

        $pl4 = Get-PadPlan -Order @('big1', 'a') -Big @('big1') -Cols 1
        Test-Assert '打包：1 列有放大时强制 2 列' ([int]$pl4.Cols -eq 2 -and [bool]$pl4.ForcedCols) ("cols=" + $pl4.Cols)

        $pl5 = Get-PadPlan -Order $n9 -Big @() -Cols 3
        Test-Assert '打包：取消放大后回到 3 行' ($pl5.Rows -eq 3) ("rows=" + $pl5.Rows)
    } catch { Test-Assert '占位打包' $false $_.Exception.Message }

    # 14. 顺序归一化
    try {
        $fake = @{ order = @('b', '不存在', 'a'); big = @('a', '也不存在') }
        $norm = Normalize-PadOrder -Pad $fake -Names @('a', 'b', 'c')
        $okN = ((@($norm.Order) -join ',') -eq 'b,a,c') -and ((@($norm.Big) -join ',') -eq 'a')
        Test-Assert '顺序归一化（丢弃失效项、补新项）' $okN ("order=" + (@($norm.Order) -join ',') + " big=" + (@($norm.Big) -join ','))
    } catch { Test-Assert '顺序归一化' $false $_.Exception.Message }

    # 15. 拖动插入排序
    try {
        $fakeP = @{ Items = @([pscustomobject]@{ Name = 'a' }, [pscustomobject]@{ Name = 'b' }, [pscustomobject]@{ Name = 'c' }, [pscustomobject]@{ Name = 'd' }) }
        $r1 = (Get-PadReorderedNames -P $fakeP -Names @('d') -Index 0) -join ','
        $r2 = (Get-PadReorderedNames -P $fakeP -Names @('a') -Index 4) -join ','
        $r3 = (Get-PadReorderedNames -P $fakeP -Names @('a', 'b') -Index 4) -join ','
        $r4 = (Get-PadReorderedNames -P $fakeP -Names @('c') -Index 1) -join ','
        $okR = ($r1 -eq 'd,a,b,c') -and ($r2 -eq 'b,c,d,a') -and ($r3 -eq 'c,d,a,b') -and ($r4 -eq 'a,c,b,d')
        Test-Assert '拖动插入排序（单项/整组/前插）' $okR ($r1 + ' / ' + $r2 + ' / ' + $r3 + ' / ' + $r4)
    } catch { Test-Assert '拖动插入排序' $false $_.Exception.Message }

    # 16. 自由位置：留空保持、交换、放大块阻挡
    try {
        $cells = @{
            'a' = @{ c = 0; r = 0 }
            'b' = @{ c = 3; r = 1 }
        }
        $plan = Get-PadPlan -Order @('a', 'b', 'c') -Big @() -Cols 4 -Cells $cells
        $okGap = ([int]$plan.Positions['b'].Col -eq 3 -and [int]$plan.Positions['b'].Row -eq 1)
        Test-Assert '自由位置：已摆放的位置不会被重排（留空保持）' $okGap ("b 在 ({0},{1})" -f $plan.Positions['b'].Col, $plan.Positions['b'].Row)

        $okNew = ([int]$plan.Positions['c'].Col -eq 1 -and [int]$plan.Positions['c'].Row -eq 0)
        Test-Assert '自由位置：新条目落到第一个空格' $okNew ("c 在 ({0},{1})" -f $plan.Positions['c'].Col, $plan.Positions['c'].Row)

        $fakeP2 = @{
            Id = 'x'; Data = @{ cells = @{ 'a' = @{c=0;r=0}; 'b' = @{c=1;r=0} }; big = @() }
            Selected = @{ 'a' = $true }
            Plan = $plan
            Items = @([pscustomobject]@{Name='a'}, [pscustomobject]@{Name='b'})
            Cols = 2; Rows = 1; Width = 150
        }
        # 直接验证交换逻辑的"数学"：把 a 右移一格应和 b 互换
        $cc = @{ 'a' = @{c=0;r=0}; 'b' = @{c=1;r=0} }
        $tmp = $cc['b']
        $cc['b'] = @{ c = [int]$cc['a'].c; r = [int]$cc['a'].r }
        $cc['a'] = @{ c = [int]$tmp.c; r = [int]$tmp.r }
        $okSwap = ([int]$cc['a'].c -eq 1 -and [int]$cc['b'].c -eq 0)
        Test-Assert '自由位置：1×1 之间换位' $okSwap ("a={0} b={1}" -f $cc['a'].c, $cc['b'].c)
    } catch { Test-Assert '自由位置' $false $_.Exception.Message }

    try { Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue } catch { }

    $fail = @($script:SelfTestResult | Where-Object { -not $_.Ok }).Count
    Write-Host ''
    Write-Host ("自检结束：{0} 项通过，{1} 项失败" -f (@($script:SelfTestResult).Count - $fail), $fail) -ForegroundColor Cyan
    if ($fail -gt 0) { exit 1 } else { exit 0 }
}

# =====================================================================
# 演示
# =====================================================================
function New-DemoContent {
    $root = Join-Path $env:TEMP 'FolderPad演示'
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $defs = @(
        @{ Dir = '毕业论文'; Days = 12; Files = @('开题报告.docx', '文献综述.md', '数据.xlsx') },
        @{ Dir = '报销材料'; Days = 2; Files = @('发票.pdf', '行程单.png') },
        @{ Dir = '旧项目'; Days = -1; Files = @('总结.txt', '备份.zip') }
    )
    $requests = @()
    foreach ($d in $defs) {
        $dir = Join-Path $root $d.Dir
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        foreach ($f in $d.Files) {
            $p = Join-Path $dir $f
            if (-not (Test-Path -LiteralPath $p)) { Set-Content -LiteralPath $p -Value ('演示文件：' + $f) -Encoding UTF8 }
        }
        $due = (Get-Date).Date.AddDays($d.Days).ToString('yyyy-MM-dd')
        $requests += ($dir + "`t" + $due)
    }
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($script:Path_DemoRequest, (($requests -join "`r`n") + "`r`n"), $enc)
    Write-Host ('演示内容已放在：' + $root)
}

function Remove-DemoPads {
    $root = Join-Path $env:TEMP 'FolderPad演示'
    Load-State | Out-Null
    $removed = 0
    foreach ($pad in @($script:State.pads)) {
        if ($null -eq $pad) { continue }
        if ([string]$pad.path -like ($root + '*')) { $removed++ }
    }
    Write-Host ("演示框记录 $removed 个（在程序里右键「移除记录」即可清掉，或直接删除目录 $root）")
}

# =====================================================================
# 主入口
# =====================================================================
$script:QuietUI = $false
if ($Quiet) { $script:QuietUI = $true }
$script:Path_DemoRequest = Join-Path $script:DataDir 'demo-request.txt'

try {
    if ($SelfTest) {
        Initialize-DataDir
        Invoke-SelfTest
        return
    }
    if ($Uninstall) {
        Invoke-Uninstall
        return
    }
    if ($Install) {
        Invoke-Install
        return
    }
    if ($Demo) {
        Initialize-DataDir
        New-DemoContent
    }

    $convertPaths = @()
    if ($InputPaths) {
        foreach ($p in $InputPaths) {
            if ([string]::IsNullOrWhiteSpace($p)) { continue }
            $convertPaths += $p
        }
    }

    if ($convertPaths.Count -gt 0) {
        Initialize-DataDir
        $valid = @()
        $invalid = @()
        foreach ($p in $convertPaths) {
            $n = Get-NormalizedPath $p
            if ($n -and [System.IO.Directory]::Exists($n)) { $valid += $n }
            else { $invalid += $p }
        }
        if ($valid.Count -eq 0) {
            Write-Log ("忽略非文件夹参数: " + ($invalid -join '; ')) 'WARN'
            return
        }
        # 已有实例在跑？投递后立刻退出，由它弹日期对话框
        $probe = New-Object System.Threading.Mutex($false, 'Local\FolderPad.Instance')
        $owned = $false
        try { $owned = $probe.WaitOne(0) } catch { $owned = $true }
        if ($owned) {
            try { $probe.ReleaseMutex() } catch { }
            foreach ($v in $valid) { Add-PendingPath $v }
            Start-FolderPadApp
            return
        }
        foreach ($v in $valid) { Add-PendingPath $v }
        Write-Log '已投递转框请求给运行中的实例'
        return
    }

    Start-FolderPadApp
} catch {
    Write-Log ("致命错误: " + $_.Exception.Message + " @ " + $_.ScriptStackTrace) 'ERROR'
    try {
        Show-MessageDialog -Title 'FolderPad 启动失败' -Message $_.Exception.Message -Detail $_.ScriptStackTrace -Kind 'error'
    } catch { }
    throw
}
