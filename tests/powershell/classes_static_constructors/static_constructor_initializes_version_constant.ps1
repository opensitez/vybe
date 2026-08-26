# vybe-test: powershell/classes_static_constructors/static_constructor_initializes_version_constant
class VersionInfo {
    static [version]$CurrentVersion
    static VersionInfo() {
        [VersionInfo]::CurrentVersion = [version]::new(7, 4, 0)
    }
}
if ([VersionInfo]::CurrentVersion.Major -ne 7 -or [VersionInfo]::CurrentVersion.Minor -ne 4) {
    Write-Host "FAIL: Static version initialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
