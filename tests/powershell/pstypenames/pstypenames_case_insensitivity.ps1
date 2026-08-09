# vybe-test: powershell/pstypenames/pstypenames_case_insensitivity
$obj = [pscustomobject]@{ K = "V" }
$types = $obj.pstypenames
if ($types -eq $null) {
    Write-Host "FAIL: case-insensitive pstypenames property access returned null"
    exit 1
}
Write-Host "PASS"
exit 0
