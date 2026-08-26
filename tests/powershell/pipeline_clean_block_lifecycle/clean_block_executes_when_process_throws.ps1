# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_executes_when_process_throws
$global:CleanRan = $false
function Test-CleanThrow {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Val)
    process {
        if ($Val -eq 2) { throw "ErrorOn2" }
    }
    clean {
        $global:CleanRan = $true
    }
}
$caught = $false
try {
    1, 2, 3 | Test-CleanThrow
} catch {
    $caught = $true
}
if (-not $caught -or -not $global:CleanRan) {
    Write-Host "FAIL: Clean block should run on terminating error"
    exit 1
}
Write-Host "PASS"
exit 0
