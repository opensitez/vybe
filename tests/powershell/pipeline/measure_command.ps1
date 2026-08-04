# vybe-test: powershell/pipeline/measure_command
$result = Measure-Command { Start-Sleep -Milliseconds 10 }
$isTimeSpan = $result -is [TimeSpan]
if ($isTimeSpan -ne $true) {
    Write-Host "FAIL: expected TimeSpan object"
    exit 1
}
Write-Host "PASS"
exit 0
