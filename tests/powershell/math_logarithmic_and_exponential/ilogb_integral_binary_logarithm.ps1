# vybe-test: powershell/math_logarithmic_and_exponential/ilogb_integral_binary_logarithm
$ilog = [math]::ILogB(1024.0)
if ($ilog -ne 10) {
    Write-Host "FAIL: ILogB(1024) expected 10, got $ilog"
    exit 1
}
Write-Host "PASS"
exit 0
