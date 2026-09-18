# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_multiple_sequential_records
# Consecutive calls to $PSCmdlet.WriteWarning emit discrete WarningRecord items in exact sequential order
function EmitSequenceOfWarnings {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteWarning("warning-alpha")
        $PSCmdlet.WriteWarning("warning-beta")
        $PSCmdlet.WriteWarning("warning-gamma")
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Continue"

try {
    $records = @(EmitSequenceOfWarnings 3>&1)

    if ($records.Count -ne 3) {
        Write-Host "FAIL: expected 3 warning records, got $($records.Count)"
        exit 1
    }

    $expectedSeq = @("warning-alpha", "warning-beta", "warning-gamma")
    for ($i = 0; $i -lt 3; $i++) {
        if ($records[$i].Message -ne $expectedSeq[$i]) {
            Write-Host "FAIL: at index ${i}, expected '$($expectedSeq[$i])', got '$($records[$i].Message)'"
            exit 1
        }
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
