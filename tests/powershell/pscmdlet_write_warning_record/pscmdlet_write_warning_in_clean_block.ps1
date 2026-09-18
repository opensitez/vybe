# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_in_clean_block
# $PSCmdlet.WriteWarning emits WarningRecord entries from within the PowerShell 7.3+ clean block
function CleanWarningEmitter {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    process {
        $InputObject
    }
    clean {
        $PSCmdlet.WriteWarning("cleanup phase detected open file handles")
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Continue"

try {
    $results = @(1 | CleanWarningEmitter 3>&1)

    # 1 data output + 1 WarningRecord from clean
    if ($results.Count -ne 2) {
        Write-Host "FAIL: expected 2 stream entries, got $($results.Count)"
        exit 1
    }

    $warningRecord = $results | Where-Object { $_ -is [System.Management.Automation.WarningRecord] }
    if ($warningRecord -eq $null -or $warningRecord.Message -ne "cleanup phase detected open file handles") {
        Write-Host "FAIL: clean block WarningRecord mismatch"
        exit 1
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
