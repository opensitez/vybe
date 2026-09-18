# vybe-test: powershell/pscmdlet_write_debug_record/pscmdlet_write_debug_in_clean_block
# $PSCmdlet.WriteDebug emits debug records from within the PowerShell 7.3+ clean block
function CleanDebugFunc {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    process {
        "work:$InputObject"
    }
    clean {
        $PSCmdlet.WriteDebug("cleanup completed successfully")
    }
}

$oldPreference = $DebugPreference
$DebugPreference = "Continue"

try {
    $results = @(1 | CleanDebugFunc *>&1)

    # 1 work string + 1 debug record from clean
    if ($results.Count -ne 2) {
        Write-Host "FAIL: expected 2 items, got $($results.Count)"
        exit 1
    }

    $debugItem = $results | Where-Object { $_ -is [System.Management.Automation.DebugRecord] }
    if ($debugItem -eq $null) {
        Write-Host "FAIL: clean block did not emit DebugRecord"
        exit 1
    }

    if ($debugItem.Message -ne "cleanup completed successfully") {
        Write-Host "FAIL: clean block message mismatch: '$($debugItem.Message)'"
        exit 1
    }
} finally {
    $DebugPreference = $oldPreference
}

Write-Host "PASS"
exit 0
