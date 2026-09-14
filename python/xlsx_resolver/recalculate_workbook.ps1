[CmdletBinding()]
param([string]$ParametersPath, [string]$CleanupParametersPath)

$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
$excelPid = $null
$ownedExcel = $false
$excelStartFileTimeUtc = $null
$statePath = $null

if (![string]::IsNullOrWhiteSpace($CleanupParametersPath)) {
    try {
        $owned = Get-Content -LiteralPath $CleanupParametersPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $p = Get-Process -Id ([int]$owned.pid) -ErrorAction SilentlyContinue
        if ($p -and $p.ProcessName -ieq 'EXCEL') {
            $ft = $p.StartTime.ToUniversalTime().ToFileTimeUtc()
            if ($ft -eq [int64]$owned.startFileTimeUtc) {
                Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
            }
        }
        exit 0
    } catch { exit 1 }
}
if ([string]::IsNullOrWhiteSpace($ParametersPath)) { throw 'ParametersPath is required' }

if (-not ('XlmergeExcelWindow' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class XlmergeExcelWindow {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@
}

function Write-State([string]$status, [string]$errorText = $null) {
    if ([string]::IsNullOrWhiteSpace($statePath)) { return }
    $state = [ordered]@{ status = $status }
    if ($null -ne $excelPid) {
        $state.pid = [int]$excelPid
        if ($null -ne $excelStartFileTimeUtc) { $state.startFileTimeUtc = [int64]$excelStartFileTimeUtc }
    }
    if ($null -ne $errorText) { $state.error = $errorText }
    $state | ConvertTo-Json -Compress | Set-Content -LiteralPath $statePath -Encoding UTF8
}

try {
    $parameters = Get-Content -LiteralPath $ParametersPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $source = [IO.Path]::GetFullPath([string]$parameters.source)
    $result = [IO.Path]::GetFullPath([string]$parameters.result)
    $statePath = [IO.Path]::GetFullPath([string]$parameters.state)
    $timeoutSeconds = [int]$parameters.timeoutSeconds
    if ($timeoutSeconds -lt 1) { throw 'timeoutSeconds must be positive' }
    if (!(Test-Path -LiteralPath $source -PathType Leaf)) { throw "Source workbook does not exist: $source" }
    $parent = Split-Path -Parent $result
    if (!(Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }

    # Excel constants are used as numbers so this script has no interop assembly
    # dependency.  AutomationSecurity=3 is msoAutomationSecurityForceDisable.
    $existingExcelPids = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $excel = New-Object -ComObject Excel.Application
    $window = [IntPtr]$excel.Hwnd
    [uint32]$windowPid = 0
    [XlmergeExcelWindow]::GetWindowThreadProcessId($window, [ref]$windowPid) | Out-Null
    if ($windowPid -eq 0) { throw 'Unable to identify the Excel process created for this calculation' }
    $excelProcess = Get-Process -Id ([int]$windowPid) -ErrorAction Stop
    if ($excelProcess.ProcessName -ine 'EXCEL') { throw 'Excel window belongs to an unexpected process' }
    if ($existingExcelPids -contains $excelProcess.Id) { throw 'Excel COM attached to a pre-existing user Excel process' }
    $excelPid = $excelProcess.Id
    $excelStartFileTimeUtc = [int64]$excelProcess.StartTime.ToUniversalTime().ToFileTimeUtc()
    $ownedExcel = $true
    Write-State 'running'

    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $excel.UserControl = $false
    # Iterative calculation can make a circular workbook appear to have a
    # stable value.  It is never safe for a conflict snapshot.
    if ($excel.Iteration) { throw 'Excel iterative calculation is enabled' }

    # UpdateLinks=0 and ReadOnly=true are positional arguments in the Excel COM
    # API.  The temporary source is never saved in place.
    $book = $excel.Workbooks.Open($source, 0, $true)
    if ($excel.Iteration) { throw 'Workbook enables iterative calculation' }
    $excel.CalculateFullRebuild()
    $deadline = [DateTime]::UtcNow.AddSeconds($timeoutSeconds)
    while ($excel.CalculationState -ne 0) { # xlDone
        if ([DateTime]::UtcNow -gt $deadline) { throw 'Excel calculation did not reach xlDone before timeout' }
        Start-Sleep -Milliseconds 100
    }
    if ([DateTime]::UtcNow -gt $deadline) { throw 'Excel calculation exceeded timeout' }
    try {
        $circular = $excel.CircularReference
        if ($null -ne $circular) { throw 'Workbook contains a circular reference' }
    } catch {
        if ($_.Exception.Message -like '*circular reference*') { throw }
        throw "Unable to verify circular references: $($_.Exception.Message)"
    }
    # Value2 is used only as a typed error probe.  The actual typed value is
    # read from Excel's serialized cache by openpyxl so dates retain the
    # workbook epoch/number format and booleans never become numeric 1/0.
    foreach ($sheet in $book.Worksheets) {
        $used = $sheet.UsedRange
        foreach ($cell in $used.Cells) {
            if ($cell.HasFormula) {
                try {
                    if ($excel.WorksheetFunction.IsError($cell.Value2)) {
                        throw "Excel returned an error value at $($sheet.Name)!$($cell.Address($false,$false))"
                    }
                } catch {
                    if ($_.Exception.Message -like '*returned an error value*') { throw }
                    throw "Unable to verify formula result at $($sheet.Name)!$($cell.Address($false,$false)): $($_.Exception.Message)"
                }
            }
        }
    }
    $book.SaveCopyAs($result)
    if (!(Test-Path -LiteralPath $result -PathType Leaf)) { throw 'Excel did not produce a result file' }
    Write-State 'completed'
    exit 0
}
catch {
    Write-State 'failed' $_.Exception.Message
    exit 1
}
finally {
    if ($null -ne $book) {
        try { $book.Close($false) } catch { }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null } catch { }
    }
    if ($null -ne $excel -and $ownedExcel) {
        try { $excel.Quit() } catch { }
    }
    if ($null -ne $excel) {
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    # Quit can leave an orphaned hidden Excel process while COM child ranges
    # are being finalized.  Ownership is exact: PID + creation timestamp.
    if ($null -ne $excelProcess -and $null -ne $excelPid) {
        try {
            $left = Get-Process -Id ([int]$excelPid) -ErrorAction SilentlyContinue
            if ($left -and $left.ProcessName -ieq 'EXCEL' -and
                $left.StartTime.ToUniversalTime().ToFileTimeUtc() -eq $excelStartFileTimeUtc) {
                Stop-Process -Id ([int]$excelPid) -Force -ErrorAction SilentlyContinue
            }
        } catch { }
    }
}
