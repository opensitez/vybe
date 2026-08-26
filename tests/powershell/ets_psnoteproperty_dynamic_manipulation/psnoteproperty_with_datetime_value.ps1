# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_with_datetime_value
$dt = [datetime]::UtcNow
$obj = [pscustomobject]@{ Timestamp = $dt }
if ($obj.Timestamp -ne $dt -or $obj.Timestamp.Year -lt 2026) {
    Write-Host "FAIL: PSNoteProperty with DateTime value failed"
    exit 1
}
Write-Host "PASS"
exit 0
