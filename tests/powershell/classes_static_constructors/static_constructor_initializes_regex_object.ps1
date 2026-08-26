# vybe-test: powershell/classes_static_constructors/static_constructor_initializes_regex_object
class RegexHolders {
    static [regex]$NumberPattern
    static RegexHolders() {
        [RegexHolders]::NumberPattern = [regex]::new("^\d+$")
    }
}
$m = [RegexHolders]::NumberPattern.IsMatch("12345")
if (-not $m) {
    Write-Host "FAIL: Static regex object initialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
