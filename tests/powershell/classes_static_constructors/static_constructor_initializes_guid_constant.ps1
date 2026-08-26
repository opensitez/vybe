# vybe-test: powershell/classes_static_constructors/static_constructor_initializes_guid_constant
class AppGuid {
    static [guid]$FixedGuid
    static AppGuid() {
        [AppGuid]::FixedGuid = [guid]::Parse("12345678-1234-1234-1234-1234567890ab")
    }
}
if ([AppGuid]::FixedGuid.ToString() -ne "12345678-1234-1234-1234-1234567890ab") {
    Write-Host "FAIL: Static GUID initialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
