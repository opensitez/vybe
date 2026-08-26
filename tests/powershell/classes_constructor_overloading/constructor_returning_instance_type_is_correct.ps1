# vybe-test: powershell/classes_constructor_overloading/constructor_returning_instance_type_is_correct
class SampleType {
    SampleType() {}
}
$s = [SampleType]::new()
if ($s -isnot [SampleType] -or $s.GetType().Name -ne "SampleType") {
    Write-Host "FAIL: Constructor return instance type check failed"
    exit 1
}
Write-Host "PASS"
exit 0
