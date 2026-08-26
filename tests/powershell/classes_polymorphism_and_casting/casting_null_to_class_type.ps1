# vybe-test: powershell/classes_polymorphism_and_casting/casting_null_to_class_type
$n = [Animal]$null
if ($n -ne $null) {
    Write-Host "FAIL: Cast null to class type must be null"
    exit 1
}
Write-Host "PASS"
exit 0
