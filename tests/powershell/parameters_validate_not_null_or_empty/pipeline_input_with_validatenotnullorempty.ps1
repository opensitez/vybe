# vybe-test: powershell/parameters_validate_not_null_or_empty/pipeline_input_with_validatenotnullorempty
function Test-NonNullPipe {
    param(
        [Parameter(ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Data
    )
    process { "DATA:$Data" }
}
$res = "ValidPayload" | Test-NonNullPipe
if ($res -ne "DATA:ValidPayload") {
    Write-Host "FAIL: Pipeline input to ValidateNotNullOrEmpty failed"
    exit 1
}
Write-Host "PASS"
exit 0
