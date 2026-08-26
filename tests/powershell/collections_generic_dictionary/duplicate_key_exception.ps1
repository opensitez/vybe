# vybe-test: powershell/collections_generic_dictionary/duplicate_key_exception
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("dup", 1)
$caught = $false
try {
    $d.Add("dup", 2)
} catch [System.ArgumentException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected ArgumentException on duplicate Add key"
    exit 1
}
Write-Host "PASS"
exit 0
