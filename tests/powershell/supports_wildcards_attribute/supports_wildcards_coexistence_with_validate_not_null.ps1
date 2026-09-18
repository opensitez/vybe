# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_coexistence_with_validate_not_null
# SupportsWildcards and ValidateNotNullOrEmpty attributes enforce validations concurrently
function SearchStrictWildcard {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [ValidateNotNullOrEmpty()]
        [string]$Pattern
    )
    process {
        return "Valid:$Pattern"
    }
}

$validCall = SearchStrictWildcard -Pattern "*.txt"
if ($validCall -ne "Valid:*.txt") {
    Write-Host "FAIL: valid call failed"
    exit 1
}

$threwEmpty = $false
try {
    SearchStrictWildcard -Pattern "" -ErrorAction Stop
} catch [System.Management.Automation.ParameterBindingException] {
    $threwEmpty = $true
}

if (-not $threwEmpty) {
    Write-Host "FAIL: ValidateNotNullOrEmpty did not throw ParameterBindingException on empty string"
    exit 1
}

Write-Host "PASS"
exit 0
