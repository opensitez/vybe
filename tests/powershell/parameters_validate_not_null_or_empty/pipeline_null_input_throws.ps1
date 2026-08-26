# vybe-test: powershell/parameters_validate_not_null_or_empty/pipeline_null_input_throws
function Test-NonNullPipe2 {
    param(
        [Parameter(ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Data
    )
    process { $Data }
}
$caught = $false
try {
    $null | Test-NonNullPipe2 -ErrorAction Stop
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception on null pipeline item"
    exit 1
}
Write-Host "PASS"
exit 0
