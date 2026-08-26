# vybe-test: powershell/classes_static_constructors/static_constructor_type_metadata_check
class MetaCheck {
    static MetaCheck() {}
}
$constructors = [MetaCheck].GetConstructors([System.Reflection.BindingFlags]"Static,NonPublic")
if ($constructors.Length -ne 1) {
    Write-Host "FAIL: Static constructor reflection metadata check failed"
    exit 1
}
Write-Host "PASS"
exit 0
