param(
    [Parameter(Mandatory = $true)][string]$Recipient,
    [Parameter(Mandatory = $true)][string]$Message
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
}
"@

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
[NativeWindow]::ShowWindow($handle, 9) | Out-Null
[NativeWindow]::SetForegroundWindow($handle) | Out-Null
Try-FocusElement $window | Out-Null
Start-Sleep -Milliseconds 400

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
    $inputX = [int]($rect.Left + ($width * 0.68))
    $bottomInset = [Math]::Max(45, [Math]::Min(70, [int]($height * 0.07)))
    $inputY = $rect.Bottom - $bottomInset
    Click-Point $inputX $inputY
}

Start-Sleep -Milliseconds 200
[System.Windows.Forms.SendKeys]::SendWait("^a")
Paste-Text $Message
Start-Sleep -Milliseconds 200
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Write-Output "SENT"
