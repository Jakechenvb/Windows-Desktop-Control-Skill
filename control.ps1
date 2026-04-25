[CmdletBinding()]
param(
    $action,
    $x = 0,
    $y = 0,
    $text = "",
    $scroll = 0
)

trap {
    Write-Error "执行错误: $_"
    exit 1
}

if ([string]::IsNullOrEmpty($action)) {
    Write-Warning "请指定 -action 参数: move/click/rightclick/doubleclick/middleclick/wheel/snapshot/type/selectall/copy/paste"
    exit 1
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$rng = [Random]::new()
$workspace = Join-Path $HOME ".openclaw\Workspace"
if (!(Test-Path $workspace)) {
    New-Item -ItemType Directory -Force -Path $workspace | Out-Null
}

function Get-Rnd {
    param($min, $max)
    return $rng.Next($min, $max + 1)
}

function Get-RndDouble {
    param($min, $max)
    return $rng.NextDouble() * ($max - $min) + $min
}

function Get-Bezier {
    param($t, $p0, $p1, $p2, $p3)
    $x = [Math]::Pow(1-$t, 3)*$p0.x + 3*[Math]::Pow(1-$t, 2)*$t*$p1.x + 3*(1-$t)*$t*$t*$p2.x + [Math]::Pow($t, 3)*$p3.x
    $y = [Math]::Pow(1-$t, 3)*$p0.y + 3*[Math]::Pow(1-$t, 2)*$t*$p1.y + 3*(1-$t)*$t*$t*$p2.y + [Math]::Pow($t, 3)*$p3.y
    return [PSCustomObject]@{ x = [int]$x; y = [int]$y }
}

function Safe-Int {
    param($value)
    $result = 0
    if (-not [int]::TryParse($value, [ref]$result)) { return 0 }
    return $result
}

function Move-Mouse-Human {
    param($tx, $ty)
    $tx = Safe-Int $tx
    $ty = Safe-Int $ty

    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $cur = [System.Windows.Forms.Cursor]::Position
    $start = @{x = $cur.X; y = $cur.Y}
    $end = @{x = $tx; y = $ty}

    $dx = $tx - $start.x
    $dy = $ty - $start.y
    $dist = [Math]::Sqrt($dx * $dx + $dy * $dy)

    if ($dist -gt 300) {
        $ratio = 0.85 + (Get-RndDouble -0.06 0.06)
        $midX = $start.x + $dx * $ratio
        $midY = $start.y + $dy * $ratio
        Move-Mouse-Human-Smooth $midX $midY
        Start-Sleep -Milliseconds (Get-Rnd 45 125)
        Move-Mouse-Human-Smooth $tx $ty
        return
    }

    Move-Mouse-Human-Smooth $tx $ty
}

function Move-Mouse-Human-Smooth {
    param($tx, $ty)

    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $cur = [System.Windows.Forms.Cursor]::Position
    $start = @{x = $cur.X; y = $cur.Y}
    $end = @{x = $tx; y = $ty}

    $dx = $tx - $start.x
    $dy = $ty - $start.y
    $dist = [Math]::Sqrt($dx * $dx + $dy * $dy)

    $baseSteps = [Math]::Max(20, [int]([Math]::Sqrt($dist) * 2.5))
    $jitterSteps = Get-RndDouble -0.25 0.25
    $steps = [Math]::Clamp([int]($baseSteps * (1 + $jitterSteps)), 20, 50)

    $jitter = Get-RndDouble 1.0 1.5
    $wave = Get-RndDouble 1.3 1.8
    $tremorStep = Get-RndDouble 0.35 0.42
    $tremorDamping = Get-RndDouble 0.84 0.88

    $cp1 = @{
        x = [Math]::Clamp($start.x + (Get-Rnd -140 140), 0, $screen.Width)
        y = [Math]::Clamp($start.y + (Get-Rnd -120 120), 0, $screen.Height)
    }
    $cp2 = @{
        x = [Math]::Clamp($end.x + (Get-Rnd -120 120), 0, $screen.Width)
        y = [Math]::Clamp($end.y + (Get-Rnd -100 100), 0, $screen.Height)
    }

    $tremorX = 0.0
    $tremorY = 0.0

    for ($i = 0; $i -le $steps; $i++) {
        $segments = 3
        $segLen = $steps / $segments
        $segIdx = [int]($i / $segLen)
        $segT = ($i % $segLen) / $segLen

        $easePower = Get-RndDouble 2.4 4.2
        $segEase = 1 - [Math]::Pow(1 - $segT, $easePower)
        $t = ($segIdx + $segEase) / $segments

        $pt = Get-Bezier $t $start $cp1 $cp2 $end
        $attenuation = 1 - [Math]::Pow($t, 2)

        $jx = Get-RndDouble (-$jitter) $jitter * $attenuation
        $jy = Get-RndDouble (-$jitter) $jitter * $attenuation
        $wx = [Math]::Sin($i * 0.65) * (Get-RndDouble 0.3 $wave) * $attenuation
        $wy = [Math]::Cos($i * 0.55) * (Get-RndDouble 0.3 $wave) * $attenuation

        $tremorX += Get-RndDouble -$tremorStep $tremorStep
        $tremorY += Get-RndDouble -$tremorStep $tremorStep
        $tremorX *= $tremorDamping
        $tremorY *= $tremorDamping
        $tremorX *= $attenuation
        $tremorY *= $attenuation

        $nx = [Math]::Clamp([int]($pt.x + $jx + $wx + $tremorX), 0, $screen.Width)
        $ny = [Math]::Clamp([int]($pt.y + $jy + $wy + $tremorY), 0, $screen.Height)
        [System.Windows.Forms.Cursor]::Position = [System.Drawing.Point]::new($nx, $ny)

        $baseDelay = 6.5 + (Get-RndDouble -1.5 1.5)
        $hesitation = 0
        if ($t -gt 0.6 -and $t -lt 0.85 -and (Get-Rnd 1 100) -le 8) {
            $hesitation = Get-Rnd 15 40
        }
        if ((Get-Rnd 1 100) -le 2) {
            Start-Sleep -Milliseconds 0
        } else {
            Start-Sleep -Milliseconds ([int]($baseDelay + $hesitation))
        }
    }

    [System.Windows.Forms.Cursor]::Position = [System.Drawing.Point]::new($tx, $ty)
    Write-Verbose "生物级移动完成: $tx , $ty"
}

if (-not ([System.Management.Automation.PSTypeName]'Win32').Type) {
    Add-Type @"
public class Win32 {
    [DllImport("user32.dll")] public static extern void mouse_event(int flags, int dx, int dy, int data, System.IntPtr extra);
    [DllImport("user32.dll")] public static extern void keybd_event(byte key, byte scan, int flags, System.IntPtr extra);
}
"@
}

function Click-Mouse {
    [Win32]::mouse_event(0x0002, 0,0,0,0)
    Start-Sleep -Milliseconds (Get-Rnd 55 125)
    [Win32]::mouse_event(0x0004, 0,0,0,0)
    Start-Sleep -Milliseconds (Get-Rnd 180 480)
}

function RightClick-Mouse {
    [Win32]::mouse_event(0x0008, 0,0,0,0)
    Start-Sleep -Milliseconds (Get-Rnd 55 125)
    [Win32]::mouse_event(0x0010, 0,0,0,0)
}

function DoubleClick-Mouse {
    Click-Mouse
    Start-Sleep -Milliseconds (Get-Rnd 210 360)
    Click-Mouse
}

function MiddleClick-Mouse {
    [Win32]::mouse_event(0x0020, 0,0,0,0)
    Start-Sleep -Milliseconds (Get-Rnd 55 125)
    [Win32]::mouse_event(0x0040, 0,0,0,0)
}

function Wheel-Mouse {
    param($delta)
    [Win32]::mouse_event(0x0800, 0, 0, $delta, 0)
}

$keyMap = @{
    'a'=65;'b'=66;'c'=67;'d'=68;'e'=69;'f'=70;'g'=71;'h'=72;'i'=73;'j'=74;
    'k'=75;'l'=76;'m'=77;'n'=78;'o'=79;'p'=80;'q'=81;'r'=82;'s'=83;'t'=84;
    'u'=85;'v'=86;'w'=87;'x'=88;'y'=89;'z'=90;
    '0'=48;'1'=49;'2'=50;'3'=51;'4'=52;'5'=53;'6'=54;'7'=55;'8'=56;'9'=57;
    ' '=32;'`'=192;'~'=192;'-'=189;'_'=189;'='=187;'+'=187;'['=219;'{'=219;
    ']'=221;'}'=221;';'=186;':'=186;'''=222;'"'=222;','=188;'<'=188;'.'=190;
    '>'=190;'/'=191;'?'=191;'\'=220;'|'=220;'!'=49;'@'=50;'#'=51;'$'=52;
    '%'=53;'^'=54;'&'=55;'*'=56;'('=57;')'=48;"`r"=13;"`n"=13;"`t"=9;"`b"=8
}
$shiftChars = '!@#$%^&*()_+{}|:"<>?~'

function Type-Text {
    param($str)
    foreach ($c in $str.ToCharArray()) {
        $char = $c.ToString()
        $lower = $char.ToLower()
        $shift = $shiftChars.Contains($char) -or [char]::IsUpper($c)

        $key = if ($keyMap.ContainsKey($char)) { $keyMap[$char] }
        elseif ($keyMap.ContainsKey($lower)) { $keyMap[$lower] }
        else { continue }

        try {
            if ($shift) { [Win32]::keybd_event(16,0,0,0); Start-Sleep -ms 18 }
            [Win32]::keybd_event([byte]$key,0,0,0)
            Start-Sleep -ms (Get-Rnd 65 170)
            [Win32]::keybd_event([byte]$key,0,2,0)
        } finally {
            if ($shift) { [Win32]::keybd_event(16,0,2,0) }
        }
        Start-Sleep -ms (Get-Rnd 35 110)
    }
}

function Select-All { try { [Win32]::keybd_event(17,0,0,0); sleep -ms 35; [Win32]::keybd_event(65,0,0,0); sleep -ms 35; [Win32]::keybd_event(65,0,2,0) } finally { [Win32]::keybd_event(17,0,2,0) } }
function Copy { try { [Win32]::keybd_event(17,0,0,0); sleep -ms 35; [Win32]::keybd_event(67,0,0,0); sleep -ms 35; [Win32]::keybd_event(67,0,2,0) } finally { [Win32]::keybd_event(17,0,2,0) } }
function Paste { try { [Win32]::keybd_event(17,0,0,0); sleep -ms 35; [Win32]::keybd_event(86,0,0,0); sleep -ms 35; [Win32]::keybd_event(86,0,2,0) } finally { [Win32]::keybd_event(17,0,2,0) } }

if ($action -eq "snapshot") {
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $bmp = $null; $g = $null
    try {
        $bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen(0,0,0,0,$screen.Size)
        $path = Join-Path $workspace "last_screen.png"
        $bmp.Save($path, [System.Drawing.Imaging.ImageFormat::Png])
        Write-Output $path
    } finally {
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
else { Write-Warning "未知指令: $action" }
