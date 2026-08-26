# vybe-test: powershell/json_nested_payload_depth/negative_depth_throws_argument_exception
$caught = $false
try {
    @{ a = 1 } | ConvertTo-Json -Depth -1
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Negative depth should throw exception"
    exit 1
}
Write-Host "PASS"
exit 0
