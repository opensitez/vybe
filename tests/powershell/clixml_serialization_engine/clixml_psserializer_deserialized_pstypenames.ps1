# vybe-test: powershell/clixml_serialization_engine/clixml_psserializer_deserialized_pstypenames
# Objects deserialized from CliXml prepend 'Deserialized.' to their type hierarchies
$orig = [pscustomobject]@{ Setting = "Optimized" }
$xml = [System.Management.Automation.PSSerializer]::Serialize($orig)
$deserialized = [System.Management.Automation.PSSerializer]::Deserialize($xml)

if ($null -eq $deserialized) {
    Write-Host "FAIL: Deserialize returned `$null"
    exit 1
}

$hasDeserializedPrefix = $false
foreach ($typeName in $deserialized.PSTypeNames) {
    if ($typeName -like "Deserialized.*") {
        $hasDeserializedPrefix = $true
        break
    }
}

if (-not $hasDeserializedPrefix) {
    Write-Host "FAIL: Deserialized prefix missing in PSTypeNames: @($($deserialized.PSTypeNames -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
