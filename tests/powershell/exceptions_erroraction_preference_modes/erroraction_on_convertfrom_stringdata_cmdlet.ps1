# vybe-test: powershell/exceptions_erroraction_preference_modes/erroraction_on_convertfrom_stringdata_cmdlet
$caught = $false
try {
    ConvertFrom-StringData "line_without_equals" -ErrorAction Stop
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: ConvertFrom-StringData -ErrorAction Stop failed"
    exit 1
}
Write-Host "PASS"
exit 0
