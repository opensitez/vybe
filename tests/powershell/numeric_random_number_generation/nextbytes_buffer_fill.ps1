# vybe-test: powershell/numeric_random_number_generation/nextbytes_buffer_fill
$rnd = [System.Random]::new(42)
[byte[]]$buf = New-Object byte[] 16
$rnd.NextBytes($buf)
$allZero = $true
foreach ($b in $buf) {
    if ($b -ne 0) { $allZero = $false; break }
}
if ($allZero) {
    Write-Host "FAIL: NextBytes filled only zeros"
    exit 1
}
Write-Host "PASS"
exit 0
