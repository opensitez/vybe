# vybe-test: powershell/using_statements/using_namespace_short_type_creation
using namespace System.Globalization

$ci = [CultureInfo]::new("en-US")
if ($ci.Name -ne "en-US") {
    Write-Host "FAIL: [CultureInfo] short creation expected 'en-US', got '$($ci.Name)'"
    exit 1
}
Write-Host "PASS"
exit 0
