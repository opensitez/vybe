# vybe-test: powershell/type_datetimeoffset_and_timezone/now_and_utcnow_properties
$now = [datetimeoffset]::Now
$utc = [datetimeoffset]::UtcNow
$diff = [math]::Abs(($now.ToUniversalTime() - $utc).TotalSeconds)
if ($diff -gt 5.0) {
    Write-Host "FAIL: Now and UtcNow diverged by more than 5s: $diff"
    exit 1
}
Write-Host "PASS"
exit 0
