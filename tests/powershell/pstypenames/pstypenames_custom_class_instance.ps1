# vybe-test: powershell/pstypenames/pstypenames_custom_class_instance
class MyPSTypeClass {
    [string]$Name = "TypedClass"
}
$inst = [MyPSTypeClass]::new()
if (-not ($inst.psobject.TypeNames -contains "MyPSTypeClass")) {
    Write-Host "FAIL: class instance TypeNames missing class name 'MyPSTypeClass'"
    exit 1
}
Write-Host "PASS"
exit 0
