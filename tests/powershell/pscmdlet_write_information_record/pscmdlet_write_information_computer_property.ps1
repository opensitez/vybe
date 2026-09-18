# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_computer_property
# InformationRecord.Computer accurately reflects the host machine name
function EmitComputerCheck {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteInformation("computer audit", @("Audit"))
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $record = (EmitComputerCheck 6>&1)[0]

    if ([string]::IsNullOrWhiteSpace($record.Computer)) {
        Write-Host "FAIL: Computer property was empty"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
