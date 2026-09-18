# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_emits_information_record_type
# $PSCmdlet.WriteInformation emits a record of type System.Management.Automation.InformationRecord onto stream 6
function EmitInfoTypeCheck {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteInformation("info payload", @("Diagnostics"))
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $captured = @(EmitInfoTypeCheck 6>&1)

    if ($captured.Count -ne 1) {
        Write-Host "FAIL: expected 1 captured record, got $($captured.Count)"
        exit 1
    }

    $typeName = $captured[0].GetType().FullName
    if ($typeName -ne "System.Management.Automation.InformationRecord") {
        Write-Host "FAIL: expected InformationRecord type, got '$typeName'"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
