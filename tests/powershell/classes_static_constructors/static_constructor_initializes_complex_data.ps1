# vybe-test: powershell/classes_static_constructors/static_constructor_initializes_complex_data
class StaticDataContainer {
    static [string[]]$DefaultTags
    static StaticDataContainer() {
        [StaticDataContainer]::DefaultTags = [string[]]@("dev", "test")
    }
}
if ([StaticDataContainer]::DefaultTags.Length -ne 2 -or [StaticDataContainer]::DefaultTags[0] -ne "dev") {
    Write-Host "FAIL: Static constructor complex data initialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
