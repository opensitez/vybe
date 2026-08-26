# vybe-test: powershell/collections_generic_dictionary/capacity_initialization
$d = [System.Collections.Generic.Dictionary[string, string]]::new(50)
if ($d.Count -ne 0) {
    Write-Host "FAIL: Initial capacity dictionary should have Count=0"
    exit 1
}
Write-Host "PASS"
exit 0
