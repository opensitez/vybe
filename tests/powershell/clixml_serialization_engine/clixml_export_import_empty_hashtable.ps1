# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_empty_hashtable
# Empty hashtables @{} round-trip through CliXml preserving type and Count == 0
$tmp = [System.IO.Path]::GetTempFileName()

@{} | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ($null -eq $restored) {
    Write-Host "FAIL: empty hashtable restored as `$null"
    exit 1
}

if ($restored.Count -ne 0) {
    Write-Host "FAIL: expected Count 0 for empty hashtable, got: $($restored.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
