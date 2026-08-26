# vybe-test: powershell/classes_constructor_overloading/constructor_with_version_parameter
class AppInfo {
    [version]$Ver
    AppInfo([string]$v) { $this.Ver = [version]::Parse($v) }
    AppInfo([version]$v) { $this.Ver = $v }
}
$a1 = [AppInfo]::new("2.1.0")
$a2 = [AppInfo]::new([version]"3.0.0")
if ($a1.Ver.Major -ne 2 -or $a2.Ver.Major -ne 3) {
    Write-Host "FAIL: Version constructor overloads failed"
    exit 1
}
Write-Host "PASS"
exit 0
