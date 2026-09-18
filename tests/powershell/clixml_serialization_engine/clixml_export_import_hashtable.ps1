# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_hashtable
# Hashtables are serialized with <DCT> tags and round-trip through CliXml
$tmp = [System.IO.Path]::GetTempFileName()
$orig = @{ Hostname = "node01.cluster"; Port = 9092; Secure = $true }

$orig | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ($restored["Hostname"] -ne "node01.cluster") {
    Write-Host "FAIL: Hostname key mismatch, got: '$($restored["Hostname"])'"
    exit 1
}

if ($restored["Port"] -ne 9092) {
    Write-Host "FAIL: Port key mismatch, got: $($restored["Port"])"
    exit 1
}

if ($restored["Secure"] -ne $true) {
    Write-Host "FAIL: Secure key mismatch, got: $($restored["Secure"])"
    exit 1
}

Write-Host "PASS"
exit 0
