# vybe-test: powershell/json/json_array
$arr = @(1, 2, 3, 4, 5)
$json = $arr | ConvertTo-Json -Compress
$parsed = $json | ConvertFrom-Json
if ($parsed.Count -ne 5) {
    Write-Host "FAIL: expected 5 elements, got $($parsed.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
