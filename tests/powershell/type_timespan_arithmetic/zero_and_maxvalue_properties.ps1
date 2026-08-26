# vybe-test: powershell/type_timespan_arithmetic/zero_and_maxvalue_properties
$zero = [timespan]::Zero
$max = [timespan]::MaxValue
$min = [timespan]::MinValue
if ($zero.Ticks -ne 0 -or $max -le $zero -or $min -ge $zero) {
    Write-Host "FAIL: TimeSpan static boundary constants"
    exit 1
}
Write-Host "PASS"
exit 0
