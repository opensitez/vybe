# vybe-test: powershell/json/json_array_roundtrip
$arr = @(1, 2, 3, 4, 5)
$json = $arr | ConvertTo-Json
$back = $json | ConvertFrom-Json
if ($back.Count -ne 5) { Write-Host "FAIL: count"; exit 1 }
for ($i = 0; $i -lt 5; $i++) {
    if ($back[$i] -ne ($i + 1)) {
        Write-Host "FAIL: [$i] expected $($i+1) got $($back[$i])"
        exit 1
    }
}
Write-Host "PASS"
exit 0
