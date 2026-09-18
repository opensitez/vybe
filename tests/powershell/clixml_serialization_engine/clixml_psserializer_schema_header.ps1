# vybe-test: powershell/clixml_serialization_engine/clixml_psserializer_schema_header
# [PSSerializer]::Serialize produces XML adhering to the PowerShell CliXml 1.1.0.1 schema
$xml = [System.Management.Automation.PSSerializer]::Serialize("sample")

if ($xml -notmatch '<Objs Version="1\.1\.0\.1"') {
    Write-Host "FAIL: CliXml schema version header missing: $xml"
    exit 1
}

if ($xml -notmatch 'xmlns="http://schemas\.microsoft\.com/powershell/2004/04"') {
    Write-Host "FAIL: CliXml XML namespace attribute missing"
    exit 1
}

Write-Host "PASS"
exit 0
