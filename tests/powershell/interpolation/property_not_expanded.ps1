# vybe-test: powershell/interpolation/property_not_expanded
# `"$obj.Name"` does NOT read the property — expansion stops at the variable and
# `.Name` stays literal text. Only `$($obj.Name)` reads it.
$obj = [PSCustomObject]@{ Name = 'Bob' }
$plain = "$obj.Name"
$wrapped = "$($obj.Name)"
if ($plain -eq 'Bob') {
    Write-Host "FAIL: bare form must not expand the property, got [$plain]"
    exit 1
}
if (-not $plain.EndsWith('.Name')) {
    Write-Host "FAIL: expected a trailing literal .Name, got [$plain]"
    exit 1
}
if ($wrapped -ne 'Bob') {
    Write-Host "FAIL: subexpression form should read the property, got [$wrapped]"
    exit 1
}
Write-Host 'PASS'
exit 0
