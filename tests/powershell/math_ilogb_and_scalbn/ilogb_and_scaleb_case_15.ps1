# vybe-test: powershell/math_ilogb_and_scalbn/ilogb_and_scaleb_case_15
$ilog = [math]::ILogB(8.0)
$scaled = [math]::ScaleB(1.0, 3)
if ($ilog -ne 3 -or $scaled -ne 8.0) { Write-Host "FAIL: ILogB / ScaleB failed"; exit 1 }
Write-Host "PASS"; exit 0
