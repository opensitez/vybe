# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_empty_array
# Exporting an empty collection to CliXml evaluates cleanly upon import
$tmp = [System.IO.Path]::GetTempFileName()

@() | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ($null -ne $restored) {
    Write-Host "FAIL: expected `$null from empty collection export, got: $restored"
    exit 1
}

Write-Host "PASS"
exit 0
