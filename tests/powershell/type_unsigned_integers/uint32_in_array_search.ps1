# vybe-test: powershell/type_unsigned_integers/uint32_in_array_search
[uint32[]]$arr = @([uint32]100, [uint32]200, [uint32]300)
$idx = [System.Array]::IndexOf($arr, [uint32]200)
if ($idx -ne 1) {
    Write-Host "FAIL: Array search for uint32 failed"
    exit 1
}
Write-Host "PASS"
exit 0
