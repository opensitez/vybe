# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_empty_string_message
# Calling $PSCmdlet.WriteWarning with an empty string produces a valid WarningRecord with empty Message
function EmitEmptyWarning {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteWarning("")
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Continue"

try {
    $records = @(EmitEmptyWarning 3>&1)

    if ($records.Count -ne 1) {
        Write-Host "FAIL: expected 1 WarningRecord, got $($records.Count)"
        exit 1
    }

    if ($records[0].Message -ne "") {
        Write-Host "FAIL: expected empty string message, got '$($records[0].Message)'"
        exit 1
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
