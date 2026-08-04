# vybe-test: powershell/types/type_decimal
[decimal]$x = 100.50
$type = $x.GetType().Name
if ($type -ne "Decimal") {
    Write-Host "FAIL: expected 'Decimal', got '$type'"
    exit 1
}
Write-Host "PASS"
exit 0
