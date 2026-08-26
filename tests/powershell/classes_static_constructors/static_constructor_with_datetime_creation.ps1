# vybe-test: powershell/classes_static_constructors/static_constructor_with_datetime_creation
class TimestampedClass {
    static [datetime]$LoadedTime
    static TimestampedClass() {
        [TimestampedClass]::LoadedTime = [datetime]::UtcNow
    }
}
if ([TimestampedClass]::LoadedTime.Year -lt 2026) {
    Write-Host "FAIL: Static datetime initialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
