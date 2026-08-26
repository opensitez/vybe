# vybe-test: powershell/dynamic_assembly_type_resolution/invalid_type_name_resolution_throws
$caught = $false
try {
    $x = [type]"NonExistent.Namespace.FakeType12345"
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception when resolving non-existent type"
    exit 1
}
Write-Host "PASS"
exit 0
