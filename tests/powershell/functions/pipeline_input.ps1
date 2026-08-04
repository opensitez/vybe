# vybe-test: powershell/functions/pipeline_input
function Double-Value {
    param(
        [Parameter(ValueFromPipeline=$true)]
        $Value
    )
    process {
        $Value * 2
    }
}
$result = 5 | Double-Value
if ($result -ne 10) {
    Write-Host "FAIL: expected 10, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
