# vybe-test: powershell/collections_generic_dictionary/keynotfoundexception_on_missing_key
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$caught = $false
try {
    $x = $d.get_Item("nonexistent")
} catch [System.Collections.Generic.KeyNotFoundException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected KeyNotFoundException on missing key access"
    exit 1
}
Write-Host "PASS"
exit 0
