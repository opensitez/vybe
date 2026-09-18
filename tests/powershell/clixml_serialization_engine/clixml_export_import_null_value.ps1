# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_null_value
# $null values are serialized as <Nil /> and deserialize back to $null
$xml = [System.Management.Automation.PSSerializer]::Serialize($null)

if ($xml -notmatch '<Nil\s*/>') {
    Write-Host "FAIL: expected <Nil /> tag for `$null, got: $xml"
    exit 1
}

$restored = [System.Management.Automation.PSSerializer]::Deserialize($xml)

if ($null -ne $restored) {
    Write-Host "FAIL: expected `$null after deserialization, got: $restored"
    exit 1
}

Write-Host "PASS"
exit 0
