# vybe-test: powershell/cmdlets/get_random
$num = Get-Random -Minimum 1 -Maximum 10
$inRange = ($num -ge 1) -and ($num -lt 10)
if ($inRange -ne $true) {
    Write-Host "FAIL: expected number in range 1-9, got $num"
    exit 1
}
Write-Host "PASS"
exit 0
