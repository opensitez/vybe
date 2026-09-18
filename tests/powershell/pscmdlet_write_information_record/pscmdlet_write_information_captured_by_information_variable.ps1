# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_captured_by_information_variable
# The -InformationVariable common parameter captures emitted InformationRecord objects without stream redirection
function EmitForVariableCapture {
    [CmdletBinding()]
    param()
    process {
        $PSCmdlet.WriteInformation("Captured via IV", @("AuditTag"))
        "worker output"
    }
}

$capturedList = $null
$out = EmitForVariableCapture -InformationVariable capturedList -InformationAction SilentlyContinue

if ($capturedList.Count -ne 1) {
    Write-Host "FAIL: expected 1 record in InformationVariable, got $($capturedList.Count)"
    exit 1
}

if ($capturedList[0].MessageData -ne "Captured via IV") {
    Write-Host "FAIL: unexpected MessageData: '$($capturedList[0].MessageData)'"
    exit 1
}

if ($capturedList[0].Tags[0] -ne "AuditTag") {
    Write-Host "FAIL: unexpected Tag: '$($capturedList[0].Tags[0])'"
    exit 1
}

Write-Host "PASS"
exit 0
