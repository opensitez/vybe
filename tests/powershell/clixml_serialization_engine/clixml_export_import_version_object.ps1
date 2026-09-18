# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_version_object
# System.Version instances round-trip through CliXml preserving all version components
$tmp = [System.IO.Path]::GetTempFileName()
$orig = [Version]"7.4.2.500"

$orig | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ($restored.Major -ne 7 -or $restored.Minor -ne 4 -or $restored.Build -ne 2 -or $restored.Revision -ne 500) {
    Write-Host "FAIL: Version components mismatch, got: $restored"
    exit 1
}

Write-Host "PASS"
exit 0
