# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_emits_warning_record_type
# $PSCmdlet.WriteWarning emits an object of type System.Management.Automation.WarningRecord onto stream 3
function EmitWarningTypeCheck {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteWarning("sample warning text")
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Continue"

try {
    $captured = @(EmitWarningTypeCheck 3>&1)

    if ($captured.Count -ne 1) {
        Write-Host "FAIL: expected 1 captured record, got $($captured.Count)"
        exit 1
    }

    $typeName = $captured[0].GetType().FullName
    if ($typeName -ne "System.Management.Automation.WarningRecord") {
        Write-Host "FAIL: expected WarningRecord type, got '$typeName'"
        exit 1
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
