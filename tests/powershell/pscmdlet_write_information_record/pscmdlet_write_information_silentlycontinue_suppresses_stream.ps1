# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_silentlycontinue_suppresses_stream
# When $InformationPreference is SilentlyContinue (default), $PSCmdlet.WriteInformation emits no output
function EmitSuppressedInfo {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteInformation("suppressed message", @("Tag"))
        "visible success data"
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "SilentlyContinue"

try {
    $results = @(EmitSuppressedInfo)

    if ($results.Count -ne 1) {
        Write-Host "FAIL: expected 1 output item, got $($results.Count)"
        exit 1
    }

    if ($results[0] -ne "visible success data") {
        Write-Host "FAIL: output mismatch under SilentlyContinue"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
