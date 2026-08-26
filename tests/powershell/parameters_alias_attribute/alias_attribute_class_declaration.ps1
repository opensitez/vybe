# vybe-test: powershell/parameters_alias_attribute/alias_attribute_class_declaration
class UserWidget {
    [Alias("Label")][string]$Title
}
$uw = [UserWidget]::new()
$uw.Title = "Widget1"
if ($uw.Title -ne "Widget1") {
    Write-Host "FAIL: Alias on class property failed"
    exit 1
}
Write-Host "PASS"
exit 0
