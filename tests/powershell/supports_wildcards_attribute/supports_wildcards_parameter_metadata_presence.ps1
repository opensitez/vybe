# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_parameter_metadata_presence
# (Get-Command).Parameters['Path'].Attributes reflects the presence of SupportsWildcardsAttribute
function Find-WildcardFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [SupportsWildcards()]
        [string]$Path
    )
    process {
        return "Target:$Path"
    }
}

$cmd = Get-Command Find-WildcardFile
$param = $cmd.Parameters["Path"]
$hasWildcardAttr = ($param.Attributes | Where-Object { $_ -is [System.Management.Automation.SupportsWildcardsAttribute] }) -ne $null

if (-not $hasWildcardAttr) {
    Write-Host "FAIL: SupportsWildcardsAttribute not found on Path parameter attributes"
    exit 1
}

Write-Host "PASS"
exit 0
