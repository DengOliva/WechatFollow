param(
    [Parameter(Mandatory = $true)][string]$Recipient,
    [Parameter(Mandatory = $true)][string]$Message,
    [string]$AttachmentPath = ""
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class NativeWindow {
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")]
    public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extraInfo);
    [DllImport("user32.dll")]
    public static extern int GetSystemMetrics(int index);
    [DllImport("user32.dll")]
    public static extern uint GetDpiForWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool SetProcessDpiAwarenessContext(IntPtr value);
}
"@

# UI Automation rectangles and mouse coordinates must use the same physical
# coordinate system on desktops with 125%/150% scaling.
try {
    [NativeWindow]::SetProcessDpiAwarenessContext([IntPtr](-4)) | Out-Null
} catch {
    # Older Windows versions may not expose per-monitor-v2 DPI awareness.
}

function Fail([string]$Text) {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    Write-Output $Text
    exit 1
}

function Paste-Text([string]$Text) {
    [System.Windows.Forms.Clipboard]::SetText($Text)
    Start-Sleep -Milliseconds 120
    [System.Windows.Forms.SendKeys]::SendWait("^v")
}

function Try-FocusElement($Element) {
    if ($null -eq $Element) { return $false }
    try {
        $Element.SetFocus()
        return $true
    } catch {
        return $false
    }
}

function Click-Point([int]$X, [int]$Y) {
    [NativeWindow]::SetCursorPos($X, $Y) | Out-Null
    Start-Sleep -Milliseconds 100
    [NativeWindow]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)
    [NativeWindow]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)
}

function Write-EnvironmentLog($Handle, $Rect) {
    try {
        $dpi = [NativeWindow]::GetDpiForWindow($Handle)
        if ($dpi -le 0) { $dpi = 96 }
        $screenWidth = [NativeWindow]::GetSystemMetrics(0)
        $screenHeight = [NativeWindow]::GetSystemMetrics(1)
        $windowWidth = $Rect.Right - $Rect.Left
        $windowHeight = $Rect.Bottom - $Rect.Top
        $line = (
            "{0:o}`tscreen={1}x{2}`tdpi={3}`tscale={4:N2}`twindow={5}x{6}" -f
            [DateTime]::Now, $screenWidth, $screenHeight, $dpi,
            ($dpi / 96.0), $windowWidth, $windowHeight
        )
        Add-Content -LiteralPath "$PSScriptRoot\uia-environment.log" -Value $line -Encoding utf8
    } catch {
        # Diagnostics must never prevent a notification.
    }
}

$processes = @(
    Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.ProcessName -in @("Weixin", "WeChat") }
)

if ($processes.Count -eq 0) {
    Fail "WECHAT_NOT_RUNNING"
}

$window = $null
foreach ($process in $processes) {
    if ($process.MainWindowHandle -ne 0) {
        $window = [System.Windows.Automation.AutomationElement]::FromHandle(
            $process.MainWindowHandle
        )
        if ($null -ne $window) { break }
    }
}

if ($null -eq $window) {
    $processIds = @($processes | Select-Object -ExpandProperty Id)
    $topWindows = [System.Windows.Automation.AutomationElement]::RootElement.FindAll(
        [System.Windows.Automation.TreeScope]::Children,
        [System.Windows.Automation.Condition]::TrueCondition
    )
    $bestArea = 0
    foreach ($candidate in $topWindows) {
        if ($processIds -contains $candidate.Current.ProcessId) {
            $candidateHandle = [IntPtr]$candidate.Current.NativeWindowHandle
            $bounds = $candidate.Current.BoundingRectangle
            $area = $bounds.Width * $bounds.Height
            if (
                $candidateHandle -ne [IntPtr]::Zero -and
                [NativeWindow]::IsWindowVisible($candidateHandle) -and
                $area -gt $bestArea
            ) {
                $window = $candidate
                $bestArea = $area
            }
        }
    }
}

if ($null -eq $window) {
    Fail "WECHAT_WINDOW_NOT_FOUND"
}

$handle = [IntPtr]$window.Current.NativeWindowHandle
if ($handle -eq [IntPtr]::Zero) {
    Fail "WECHAT_WINDOW_NOT_FOUND"
}
# Maximize WeChat so narrow RDP resolutions expose the largest possible editor.
[NativeWindow]::ShowWindow($handle, 3) | Out-Null
[NativeWindow]::SetForegroundWindow($handle) | Out-Null
Try-FocusElement $window | Out-Null
Start-Sleep -Milliseconds 600

$initialRect = New-Object NativeWindow+RECT
if ([NativeWindow]::GetWindowRect($handle, [ref]$initialRect)) {
    Write-EnvironmentLog $handle $initialRect
}

$editCondition = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
    [System.Windows.Automation.ControlType]::Edit
)
$edits = $window.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    $editCondition
)

$searchBox = $null
foreach ($edit in $edits) {
    $name = $edit.Current.Name
    $automationId = $edit.Current.AutomationId
    if ($name -match "Search" -or $automationId -match "search") {
        $searchBox = $edit
        break
    }
}

if ($null -ne $searchBox) {
    $searchFocused = Try-FocusElement $searchBox
}
if (-not $searchFocused) {
    # Newer WeChat versions render parts of the interface as custom controls.
    [System.Windows.Forms.SendKeys]::SendWait("^f")
}

Start-Sleep -Milliseconds 200
[System.Windows.Forms.SendKeys]::SendWait("^a")
Paste-Text $Recipient
Start-Sleep -Milliseconds 900
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Milliseconds 700

$edits = $window.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    $editCondition
)
$messageBox = $null
for ($i = $edits.Count - 1; $i -ge 0; $i--) {
    $candidate = $edits.Item($i)
    if ($null -eq $searchBox -or -not $candidate.Equals($searchBox)) {
        $messageBox = $candidate
        break
    }
}

if ($null -ne $messageBox) {
    $messageFocused = Try-FocusElement $messageBox
}
if (-not $messageFocused) {
    # Resolve the real foreground window again after opening the conversation.
    # The original UIA element can become stale or refer to a host window.
    $foreground = [NativeWindow]::GetForegroundWindow()
    $rect = New-Object NativeWindow+RECT
    $hasRect = (
        $foreground -ne [IntPtr]::Zero -and
        [NativeWindow]::GetWindowRect($foreground, [ref]$rect)
    )
    if (-not $hasRect) {
        Fail "CHAT_INPUT_NOT_FOUND"
    }
    $width = $rect.Right - $rect.Left
    $height = $rect.Bottom - $rect.Top
    if ($width -lt 500 -or $height -lt 400) {
        Fail "CHAT_INPUT_NOT_FOUND"
    }

    # Click relative to the bottom edge, not a percentage from the top.
    # When the editor is collapsed, a height-based percentage can land on
    # the emoji/file toolbar instead of the text area.
    $dpi = [NativeWindow]::GetDpiForWindow($foreground)
    if ($dpi -le 0) { $dpi = 96 }
    $scale = $dpi / 96.0
    $inputX = [int]($rect.Left + ($width * 0.68))
    $minimumInset = [int](45 * $scale)
    $maximumInset = [int](78 * $scale)
    $bottomInset = [Math]::Max(
        $minimumInset,
        [Math]::Min($maximumInset, [int]($height * 0.08))
    )
    $inputY = $rect.Bottom - $bottomInset

    # Keep the fallback click inside the current virtual desktop, including RDP.
    $virtualLeft = [NativeWindow]::GetSystemMetrics(76)
    $virtualTop = [NativeWindow]::GetSystemMetrics(77)
    $virtualRight = $virtualLeft + [NativeWindow]::GetSystemMetrics(78) - 1
    $virtualBottom = $virtualTop + [NativeWindow]::GetSystemMetrics(79) - 1
    $inputX = [Math]::Max($virtualLeft, [Math]::Min($virtualRight, $inputX))
    $inputY = [Math]::Max($virtualTop, [Math]::Min($virtualBottom, $inputY))
    Click-Point $inputX $inputY
}

Start-Sleep -Milliseconds 200
[System.Windows.Forms.SendKeys]::SendWait("^a")
Paste-Text $Message
Start-Sleep -Milliseconds 200
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

if ($AttachmentPath) {
    if (-not (Test-Path -LiteralPath $AttachmentPath -PathType Leaf)) {
        Fail "ATTACHMENT_NOT_FOUND"
    }
    $resolvedAttachment = (Resolve-Path -LiteralPath $AttachmentPath).Path
    $files = New-Object System.Collections.Specialized.StringCollection
    [void]$files.Add($resolvedAttachment)
    [System.Windows.Forms.Clipboard]::SetFileDropList($files)
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait("^v")
    Start-Sleep -Milliseconds 1500
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep -Milliseconds 500
}

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
if ($AttachmentPath) {
    Write-Output "SENT_WITH_ATTACHMENT"
} else {
    Write-Output "SENT"
}
