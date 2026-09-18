# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_captured_by_warning_variable
# The -WarningVariable common parameter collects WarningRecord objects emitted during command execution
function CollectWarningTarget {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteWarning("collected warning for variable")
        "return data"
    }
}

$capturedWarnings = $null
$out = CollectWarningTarget -WarningVariable capturedWarnings -WarningAction SilentlyContinue

if ($capturedWarnings.Count -ne 1) {
    Write-Host "FAIL: expected 1 warning record in variable, got $($capturedWarnings.Count)"
    exit 1
}

if ($capturedWarnings[0].Message -ne "collected warning for variable") {
    Write-Host "FAIL: WarningVariable message mismatch: '$($capturedWarnings[0].Message)'"
    exit 1
}

Write-Host "PASS"
exit 0
