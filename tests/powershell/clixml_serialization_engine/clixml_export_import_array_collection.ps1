# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_array_collection
# Heterogeneous arrays serialize element-by-element and restore all elements in order
$tmp = [System.IO.Path]::GetTempFileName()
$orig = @("first_string", 42, $true)

$orig | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ($restored.Count -ne 3) {
    Write-Host "FAIL: expected 3 elements in restored array, got $($restored.Count)"
    exit 1
}

if ($restored[0] -ne "first_string" -or $restored[1] -ne 42 -or $restored[2] -ne $true) {
    Write-Host "FAIL: elements mismatch: @($($restored -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
