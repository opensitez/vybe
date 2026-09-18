# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_empty_string
# An empty string serializes to an empty string tag and restores as [string]::Empty
$tmp = [System.IO.Path]::GetTempFileName()

"" | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ($null -eq $restored) {
    Write-Host "FAIL: empty string deserialized as `$null"
    exit 1
}

if ($restored -ne "") {
    Write-Host "FAIL: expected empty string, got: '$restored'"
    exit 1
}

Write-Host "PASS"
exit 0
