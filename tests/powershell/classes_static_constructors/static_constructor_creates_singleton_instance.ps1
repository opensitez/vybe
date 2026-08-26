# vybe-test: powershell/classes_static_constructors/static_constructor_creates_singleton_instance
class Singleton {
    static [Singleton]$Instance
    [string]$Status
    static Singleton() {
        [Singleton]::Instance = [Singleton]::new()
        [Singleton]::Instance.Status = "Ready"
    }
    Singleton() {}
}
if ([Singleton]::Instance.Status -ne "Ready") {
    Write-Host "FAIL: Singleton creation in static constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
