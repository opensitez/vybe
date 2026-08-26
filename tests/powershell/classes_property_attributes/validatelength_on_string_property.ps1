# vybe-test: powershell/classes_property_attributes/validatelength_on_string_property
class PasswordBox {
    [ValidateLength(6, 20)][string]$Password
}
$pb = [PasswordBox]::new()
$pb.Password = "secret123"
$caught = $false
try {
    $pb.Password = "123"
} catch {
    $caught = $true
}
if ($pb.Password -ne "secret123" -or -not $caught) {
    Write-Host "FAIL: ValidateLength on property failed"
    exit 1
}
Write-Host "PASS"
exit 0
