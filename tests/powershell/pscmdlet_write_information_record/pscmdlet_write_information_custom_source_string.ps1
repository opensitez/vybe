# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_custom_source_string
# Constructing an InformationRecord with a custom source identifier preserves the Source property value
function EmitCustomSource {
    [CmdletBinding()]
    param()
    process {
        $rec = [System.Management.Automation.InformationRecord]::new(
            "Audit message",
            "AuthenticationSubsystem.Kerberos"
        )
        $PSCmdlet.WriteInformation($rec)
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $record = (EmitCustomSource 6>&1)[0]

    if ($record.Source -ne "AuthenticationSubsystem.Kerberos") {
        Write-Host "FAIL: expected custom source 'AuthenticationSubsystem.Kerberos', got '$($record.Source)'"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
