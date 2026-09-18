# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_multiline_string
# Multiline strings containing linebreaks serialize safely and preserve linebreaks upon import
$tmp = [System.IO.Path]::GetTempFileName()
$orig = "HeaderLine`nSecondLine`n`nFourthLineWithTrailing`n"

$orig | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ($restored -ne $orig) {
    Write-Host "FAIL: multiline string corrupted during round-trip"
    exit 1
}

Write-Host "PASS"
exit 0
