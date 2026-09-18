# vybe-test: powershell/pscmdlet_write_warning_record/pscmdlet_write_warning_does_not_pollute_success_stream_without_redirection
# Without stream redirection (3>&1), $PSCmdlet.WriteWarning does not pollute variables capturing success output
function CleanDataGenerator {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteWarning("diagnostic warning emitted")
        "clean payload content"
    }
}

$oldPreference = $WarningPreference
$WarningPreference = "Continue"

try {
    $result = CleanDataGenerator

    if ($result -ne "clean payload content") {
        Write-Host "FAIL: success stream was corrupted by warning record: '$result'"
        exit 1
    }
} finally {
    $WarningPreference = $oldPreference
}

Write-Host "PASS"
exit 0
