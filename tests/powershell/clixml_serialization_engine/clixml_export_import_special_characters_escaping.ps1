# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_special_characters_escaping
# Special XML characters like &, <, >, ', and " are safely escaped and restored
$tmp = [System.IO.Path]::GetTempFileName()
$orig = "XML special: & <tag> 'quoted' `"double`" && --"

$orig | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ($restored -ne $orig) {
    Write-Host "FAIL: XML escaping corrupted string, got: '$restored'"
    exit 1
}

Write-Host "PASS"
exit 0
