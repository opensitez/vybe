# vybe-test: powershell/classes_property_attributes/validatepattern_regex_on_property
class EmailHolder {
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')][string]$Email
}
$eh = [EmailHolder]::new()
$eh.Email = "alice@example.com"
$caught = $false
try {
    $eh.Email = "not-an-email"
} catch {
    $caught = $true
}
if ($eh.Email -ne "alice@example.com" -or -not $caught) {
    Write-Host "FAIL: ValidatePattern on property failed"
    exit 1
}
Write-Host "PASS"
exit 0
