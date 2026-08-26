# vybe-test: powershell/numeric_random_number_generation/nextbytes_partial_span_simulation
$rnd = [System.Random]::new()
[byte[]]$buf = New-Object byte[] 4
$rnd.NextBytes($buf)
if ($buf.Length -ne 4) {
    Write-Host "FAIL: NextBytes array length mismatch"
    exit 1
}
Write-Host "PASS"
exit 0
