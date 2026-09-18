# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_in_begin_block
# $PSCmdlet.WriteWarning emits WarningRecord entries from within the function begin block
function BeginWarningEmitter {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    begin {
        $PSCmdlet.WriteWarning("Configuration file missing; defaulting values")
    }
    process {
        $InputObject * 3
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Continue"

try {
    $results = @(1..2 | BeginWarningEmitter 3>&1)

    # 1 WarningRecord from begin + 2 pipeline outputs
    if ($results.Count -ne 3) {
        Write-Host "FAIL: expected 3 stream entries, got $($results.Count)"
        exit 1
    }

    if ($results[0].GetType().Name -ne "WarningRecord") {
        Write-Host "FAIL: first entry was not WarningRecord"
        exit 1
    }

    if ($results[0].Message -ne "Configuration file missing; defaulting values") {
        Write-Host "FAIL: begin warning message mismatch"
        exit 1
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
