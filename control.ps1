[CmdletBinding()]
param(
    $action,
    $x = 0,
    $y = 0,
    $text = "",
    $scroll = 0
)

# 顶级错误捕获（输出干净，不崩溃）
trap {
    Write-Error "执行错误: $_"
    exit 1
}

# 1. 空 action 校验（友好提示）
if ([string]::IsNullOrEmpty($action)) {
    Write-Warning "请指定 -action 参数: move/click/rightclick/doubleclick/middleclick/wheel/snapshot/type/selectall/copy/paste"
    exit 1
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 全局随机
$rng = [Random]::new()

# 真人行为参数
$steps = 36
$jitter = 1.5
$wave = 1.8
$moveDelay = 8

# 工作目录
$workspace = Join-Path $HOME ".openclaw\Workspace"
if (!(Test-Path $workspace)) {
    New-Item -ItemType Directory -Force -Path $workspace | Out-Null
}

# 随机整数（含上下限）
function Get-Rnd {
    param($min, $max)
    return $rng.Next($min, $max + 1)
}

# 随机浮点数
function Get-RndDouble {
    param($min, $max)
    return $rng.NextDouble() * ($max - $min) + $min
}

# 贝塞尔曲线
function Get-Bezier {
    param($t, $p0, $p1, $p2, $p3)
    $x = (1-$t)**3*$p0.x + 3*(1-$t)**2*$t*$p1.x + 3*(1-$t)*$t*$t*$p2.x + $t**3*$p3.x
    $y = (1-$t)**3*$p0.y + 3*(1-$t)**2*$t*$p1.y + 3*(1-$t)*$t*$t*$p2.y + $t**3*$p3.y
    return [PSCustomObject]@{ x = [int]$x; y = [int]$y }
}

# 缓动函数：EaseOutCubic
function Get-Easing {
    param($t)
    return 1 - [Math]::Pow(1 - $t, 3)
}

# 3. 防御性坐标解析（安全转换，非数字不崩溃）
function Safe-Int {
    param($value)
    $result = 0
    if (-not [int]::TryParse($value, [ref]$result)) { return 0 }
    return $result
}

# 真人鼠标移动
function Move-Mouse-Human {
    param($tx, $ty)
    $tx = Safe-Int $tx
    $ty = Safe-Int $ty

    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $cur = [System.Windows.Forms.Cursor]::Position
    $start = @{x = $cur.X; y = $cur.Y}
    $end = @{x = $tx; y = $ty}

    $cp1 = @{
        x = [Math]::Clamp($start.x + Get-Rnd -160 160, 0, $screen.Width)
        y = [Math]::Clamp($start.y + Get-Rnd -140 140, 0, $screen.Height)
    }
    $cp2 = @{
        x = [Math]::Clamp($end.x + Get-Rnd -140 140, 0, $screen.Width)
        y = [Math]::Clamp($end.y + Get-Rnd -120 120, 0, $screen.Height)
    }

    for ($i = 0; $i -le $steps; $i++) {
        $t = $i / $steps
        $t = Get-Easing $t

        $pt = Get-Bezier $t $start $cp1 $cp2 $end
        
        $attenuation = 1 - [Math]::Pow($t, 2)
        $jx = Get-RndDouble (-$jitter) $jitter * $attenuation
        $jy = Get-RndDouble (-$jitter) $jitter * $attenuation
        $wx = [Math]::Sin($i * 0.75) * (Get-RndDouble 0.4 $wave) * $attenuation
        $wy = [Math]::Cos($i * 0.6) * (Get-RndDouble 0.4 $wave) * $attenuation

        $nx = [Math]::Clamp([int]($pt.x + $jx + $wx), 0, $screen.Width)
        $ny = [Math]::Clamp([int]($pt.y + $jy + $wy), 0, $screen.Height)

        [System.Windows.Forms.Cursor]::Position = [System.Drawing.Point]::new($nx, $ny)
        Start-Sleep -Milliseconds $moveDelay
    }

    [System.Windows.Forms.Cursor]::Position = [System.Drawing.Point]::new($tx, $ty)
    Write-Verbose "鼠标移动完成: $tx , $ty"
}

# 避免重复定义 Win32
if (-not ([System.Management.Automation.PSTypeName]'Win32').Type) {
    Add-Type @"
public class Win32 {
    [DllImport("user32.dll")]
    public static extern void mouse_event(int flags, int dx, int dy, int data, System.IntPtr extra);
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte key, byte scan, int flags, System.IntPtr extra);
}
"@
}

# 单击
function Click-Mouse {
    [Win32]::mouse_event(0x0002, 0,0,0,0)
    Start-Sleep -Milliseconds (Get-Rnd 60 130)
    [Win32]::mouse_event(0x0004, 0,0,0,0)
    Start-Sleep -Milliseconds (Get-Rnd 200 500)
    Write-Verbose "左键单击"
}

# 4. 右键点击
function RightClick-Mouse {
    [Win32]::mouse_event(0x0008, 0,0,0,0)
    Start-Sleep -Milliseconds (Get-Rnd 60 120)
    [Win32]::mouse_event(0x0010, 0,0,0,0)
    Write-Verbose "右键单击"
}

# 双击
function DoubleClick-Mouse {
    Click-Mouse
    Start-Sleep -Milliseconds (Get-Rnd 220 380)
    Click-Mouse
    Write-Verbose "左键双击"
}

# 中键点击
function MiddleClick-Mouse {
    [Win32]::mouse_event(0x0020, 0,0,0,0)
    Start-Sleep -Milliseconds (Get-Rnd 60 120)
    [Win32]::mouse_event(0x0040, 0,0,0,0)
    Write-Verbose "中键单击"
}

# 鼠标滚轮
function Wheel-Mouse {
    param($delta)
    [Win32]::mouse_event(0x0800, 0, 0, $delta, 0)
    Write-Verbose "滚轮滚动: $delta"
}

# 完整键码映射（2. 增加 LF 换行支持）
$keyMap = @{
    'a'=65; 'b'=66; 'c'=67; 'd'=68; 'e'=69; 'f'=70; 'g'=71; 'h'=72; 'i'=73; 'j'=74;
    'k'=75; 'l'=76; 'm'=77; 'n'=78; 'o'=79; 'p'=80; 'q'=81; 'r'=82; 's'=83; 't'=84;
    'u'=85; 'v'=86; 'w'=87; 'x'=88; 'y'=89; 'z'=90;
    '0'=48; '1'=49; '2'=50; '3'=51; '4'=52; '5'=53; '6'=54; '7'=55; '8'=56; '9'=57;
    ' '=32; '`'=192; '~'=192;
    '-'=189; '_'=189;
    '='=187; '+'=187;
    '['=219; '{'=219;
    ']'=221; '}'=221;
    ';'=186; ':'=186;
    ''''=222; '"'=222;
    ','=188; '<'=188;
    '.'=190; '>'=190;
    '/'=191; '?'=191;
    '\'=220; '|'=220;
    '!'=49; '@'=50; '#'=51; '$'=52; '%'=53; '^'=54; '&'=55; '*'=56; '('=57; ')'=48;
    "`r"=13; "`n"=13; "`t"=9; "`b"=8
}

$shiftChars = '!@#$%^&*()_+{}|:"<>?~'

# 真人打字
function Type-Text {
    param($str)
    foreach ($c in $str.ToCharArray()) {
        $char = $c.ToString()
        $lower = $char.ToLower()
        $shift = $shiftChars.Contains($char) -or [char]::IsUpper($c)

        if ($keyMap.ContainsKey($char)) {
            $key = $keyMap[$char]
        } elseif ($keyMap.ContainsKey($lower)) {
            $key = $keyMap[$lower]
        } else {
            continue
        }

        try {
            if ($shift) { [Win32]::keybd_event(16,0,0,0); Start-Sleep -ms 20 }
            [Win32]::keybd_event([byte]$key,0,0,0)
            Start-Sleep -ms (Get-Rnd 70 180)
            [Win32]::keybd_event([byte]$key,0,2,0)
            Write-Verbose "输入字符: $char"
        }
        finally {
            if ($shift) { [Win32]::keybd_event(16,0,2,0) }
        }
        Start-Sleep -ms (Get-Rnd 40 120)
    }
}

# 全选
function Select-All {
    try { [Win32]::keybd_event(17,0,0,0); Start-Sleep -ms 40; [Win32]::keybd_event(65,0,0,0); Start-Sleep -ms 40; [Win32]::keybd_event(65,0,2,0) }
    finally { [Win32]::keybd_event(17,0,2,0) }
    Write-Verbose "全选"
}

# 复制
function Copy {
    try { [Win32]::keybd_event(17,0,0,0); Start-Sleep -ms 40; [Win32]::keybd_event(67,0,0,0); Start-Sleep -ms 40; [Win32]::keybd_event(67,0,2,0) }
    finally { [Win32]::keybd_event(17,0,2,0) }
    Write-Verbose "复制"
}

# 粘贴
function Paste {
    try { [Win32]::keybd_event(17,0,0,0); Start-Sleep -ms 40; [Win32]::keybd_event(86,0,0,0); Start-Sleep -ms 40; [Win32]::keybd_event(86,0,2,0) }
    finally { [Win32]::keybd_event(17,0,2,0) }
    Write-Verbose "粘贴"
}

# 指令执行
if ($action -eq "snapshot") {
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $bmp = $null
    $g = $null
    try {
        $bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen(0,0,0,0,$screen.Size)
        $path = Join-Path $workspace "last_screen.png"
        $bmp.Save($path, [System.Drawing.Imaging.ImageFormat::Png])
        Write-Output $path
    }
    finally {
        if ($g) { $g.Dispose(); $g = $null }
        if ($bmp) { $bmp.Dispose(); $bmp = $null }
    }
}
elseif ($action -eq "move") { Move-Mouse-Human $x $y }
elseif ($action -eq "click") { Move-Mouse-Human $x $y; Click-Mouse }
elseif ($action -eq "rightclick") { Move-Mouse-Human $x $y; RightClick-Mouse }
elseif ($action -eq "doubleclick") { Move-Mouse-Human $x $y; DoubleClick-Mouse }
elseif ($action -eq "middleclick") { Move-Mouse-Human $x $y; MiddleClick-Mouse }
elseif ($action -eq "wheel") { Wheel-Mouse $scroll }
elseif ($action -eq "selectall") { Select-All }
elseif ($action -eq "copy") { Copy }
elseif ($action -eq "paste") { Paste }
elseif ($action -eq "type") { Type-Text $text }
else {
    Write-Warning "未知指令: $action"
}