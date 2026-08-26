# vybe-test: powershell/type_int128_and_uint128_arithmetic/int128_sign_method
$neg = [System.Int128]::Parse("-500")
$pos = [System.Int128]::Parse("500")
$zero = [System.Int128]::Zero
if ([System.Int128]::Sign($neg) -ne -1 -or [System.Int128]::Sign($pos) -ne 1 -or [System.Int128]::Sign($zero) -ne 0) {
    Write-Host "FAIL: Int128 Sign failed"; exit 1
}
Write-Host "PASS"; exit 0
