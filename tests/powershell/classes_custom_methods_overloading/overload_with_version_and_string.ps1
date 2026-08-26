# vybe-test: powershell/classes_custom_methods_overloading/overload_with_version_and_string
class VersionChecker {
    [bool]IsAtLeastV2([string]$v) { return ([version]::Parse($v).Major -ge 2) }
    [bool]IsAtLeastV2([version]$v) { return ($v.Major -ge 2) }
}
$vc = [VersionChecker]::new()
if (-not $vc.IsAtLeastV2("2.1.0") -or $vc.IsAtLeastV2([version]"1.9.0")) {
    Write-Host "FAIL: VersionChecker overload failed"
    exit 1
}
Write-Host "PASS"
exit 0
